import {
	App,
	Editor,
	EditorPosition,
	EditorSuggest,
	EditorSuggestContext,
	EditorSuggestTriggerInfo,
	FuzzySuggestModal,
	MarkdownPostProcessorContext,
	Modal,
	Notice,
	Plugin,
	TFile,
} from 'obsidian';
import * as XLSX from 'xlsx';

// Sentinel + alias for the "create new file" suggestion entry.
const CREATE_NEW = Symbol('CREATE_NEW');
type CreateNewItem = typeof CREATE_NEW;
type FileOrNew = TFile | CreateNewItem;

// Subset of the SheetJS cell-style shape we actually read. The upstream
// `XLSX.CellObject.s` is typed as `any`, so we type the parameter locally
// to keep the renderer honest.
interface CellStyle {
	fill?: { fgColor?: { rgb?: string } };
	font?: { color?: { rgb?: string }; bold?: boolean; italic?: boolean };
}

// Obsidian's public `App` type doesn't expose the desktop file-open
// helpers; the live instance has them. Add them as optional extras on the
// narrow local view so the typeof-checks still narrow correctly.
type AppWithFileOpen = App & {
	openWithDefaultApp?: (path: string) => Promise<void>;
	showInFolder?: (path: string) => void;
	revealInFolder?: (file: TFile) => void;
};

interface SpreadsheetRange {
	sheet?: string;
	startCell: string;
	endCell: string;
}

interface RenderOptions {
	mode: 'stacked' | 'tabbed';
	formatting: boolean;
}

interface ParsedSpec {
	ranges: SpreadsheetRange[];
	options: RenderOptions;
}

interface WorksheetHandle {
	workbook: XLSX.WorkBook;
	file: TFile;
	filePath: string;
	filename: string;
	el: HTMLElement; // post-processor container — used to find the section in the note
	ctx: MarkdownPostProcessorContext;
	spec: ParsedSpec; // original spec — kept so we can preserve options on rewrite
}

export default class SpreadsheetSyncPlugin extends Plugin {
	async onload() {
		this.registerMarkdownCodeBlockProcessor('spreadsheet', async (source, el, ctx) => {
			try {
				await this.renderSpreadsheetBlock(source, el, ctx);
			} catch (error) {
				el.createEl('div', { text: `Error: ${error.message}`, cls: 'spreadsheet-error' });
			}
		});

		// Command palette entry — same picker as the slash menu.
		this.addCommand({
			id: 'insert-spreadsheet-block',
			name: 'Insert spreadsheet block',
			editorCallback: (editor: Editor) => this.openXlsxFilePicker(editor),
		});

		// Slash menu — fires when the user types `/xlsx...` in a note.
		this.registerEditorSuggest(new XlsxFileSuggest(this.app, this));
	}

	private parseRange(rangeStr: string): SpreadsheetRange {
		const match = rangeStr.match(/(?:([^!]+)!)?([A-Z]+\d+):([A-Z]+\d+)/);
		if (!match) {
			throw new Error('Invalid range format. Expected format: [SheetName!]A1:B2');
		}
		return {
			sheet: match[1]?.trim(),
			startCell: match[2],
			endCell: match[3],
		};
	}

	private parseSpec(innerArg: string): ParsedSpec {
		// Split on first ';' — ranges part + options part.
		const semiIdx = innerArg.indexOf(';');
		const rangesPart = (semiIdx >= 0 ? innerArg.substring(0, semiIdx) : innerArg).trim();
		const optionsPart = (semiIdx >= 0 ? innerArg.substring(semiIdx + 1) : '').trim();

		// Ranges separated by ','.
		const rangeStrs = rangesPart.split(',').map(s => s.trim()).filter(Boolean);
		if (rangeStrs.length === 0) {
			throw new Error('No ranges provided');
		}
		const ranges = rangeStrs.map(s => this.parseRange(s));

		// Options: key=value pairs, ',' separated (mirrors the ranges separator).
		const options: RenderOptions = { mode: 'stacked', formatting: true };
		if (optionsPart) {
			for (const opt of optionsPart.split(',').map(s => s.trim()).filter(Boolean)) {
				const eqIdx = opt.indexOf('=');
				if (eqIdx < 0) continue;
				const key = opt.substring(0, eqIdx).trim().toLowerCase();
				const value = opt.substring(eqIdx + 1).trim().toLowerCase();
				if (key === 'mode') {
					if (value === 'tabbed' || value === 'stacked') {
						options.mode = value;
					}
				} else if (key === 'formatting') {
					options.formatting = value !== 'off';
				}
			}
		}

		return { ranges, options };
	}

	private async renderSpreadsheetBlock(source: string, el: HTMLElement, ctx: MarkdownPostProcessorContext) {
		const match = source.trim().match(/^(.+?)\((.+?)\)$/);
		if (!match) {
			throw new Error('Invalid format. Expected: filename.xlsx(A1:B2) or filename.xlsx(A1:B2,C1:D5; mode=tabbed)');
		}

		const [, filename, innerArg] = match;
		const spec = this.parseSpec(innerArg);

		// Resolve the file path. Obsidian mobile sandbox can't use Node's `path` module,
		// so resolve with string ops. `sourcePath` always uses forward slashes.
		const notePath = ctx.sourcePath;
		const lastSlash = notePath.lastIndexOf('/');
		const noteDir = lastSlash >= 0 ? notePath.substring(0, lastSlash) : '';
		const filePath = noteDir ? `${noteDir}/${filename}` : filename;

		const abstractFile = this.app.vault.getAbstractFileByPath(filePath);
		if (!abstractFile || !(abstractFile instanceof TFile)) {
			throw new Error(`File not found or is not a valid file: ${filename} (looking in ${filePath})`);
		}
		const file = abstractFile;

		const arrayBuffer = await this.app.vault.readBinary(file);
		const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });

		const handle: WorksheetHandle = { workbook, file, filePath, filename, el, ctx, spec };

		// Outer block
		const block = el.createEl('div', { cls: 'spreadsheet-block' });

		// Top toolbar first — must be created before the body so it renders
		// at the top of the block (DOM order = visual order).
		this.renderTopToolbar(block, handle);

		// Body container — cleared and re-rendered by the "Load all sheets"
		// button so we can swap content without losing the outer toolbar.
		const body = block.createEl('div', { cls: 'spreadsheet-body' });

		if (spec.options.mode === 'tabbed') {
			this.renderTabbedBlock(body, handle, spec.ranges, spec.options);
		} else {
			for (const range of spec.ranges) {
				this.renderSubBlock(body, handle, range, spec.options);
			}
		}
	}

	private renderTopToolbar(parent: HTMLElement, handle: WorksheetHandle) {
		const toolbar = parent.createEl('div', { cls: 'spreadsheet-toolbar' });
		const label = toolbar.createEl('span', {
			cls: 'spreadsheet-label',
			text: handle.filename,
		});
		label.title = handle.filePath;

		const actions = toolbar.createEl('span', { cls: 'spreadsheet-actions' });

		// Always show — even for single-sheet workbooks the action is still
		// useful as "re-read file and re-render", and showing it consistently
		// avoids the surprise of it appearing/disappearing as sheets change.
		const loadAllLink = actions.createEl('a', {
			cls: 'spreadsheet-load-all',
			text: 'Load all sheets',
			href: '#',
		});
		loadAllLink.title = 'Rewrite the code block to render every sheet in the workbook';
		loadAllLink.addEventListener('click', async (evt) => {
			evt.preventDefault();
			const body = parent.querySelector('.spreadsheet-body') as HTMLElement | null;
			if (body) await this.loadAllSheets(body, handle);
		});

		const openLink = actions.createEl('a', {
			cls: 'spreadsheet-open',
			text: 'Open in Excel',
			href: '#',
		});
		openLink.addEventListener('click', async (evt) => {
			evt.preventDefault();
			const app: AppWithFileOpen = this.app;
			try {
				if (typeof app.openWithDefaultApp === 'function') {
					await app.openWithDefaultApp(handle.file.path);
				} else if (typeof app.showInFolder === 'function') {
					app.showInFolder(handle.file.path);
				} else if (typeof app.revealInFolder === 'function') {
					app.revealInFolder(handle.file);
				} else {
					throw new Error('No file-open API available in this Obsidian version');
				}
			} catch (err) {
				console.error('SpreadsheetSync: openWithDefaultApp failed', err);
				const errDiv = parent.createEl('div', {
					text: `Could not open ${handle.filename}: ${err.message}. File path: ${handle.filePath}`,
					cls: 'spreadsheet-error',
				});
				setTimeout(() => errDiv.remove(), 8000);
			}
		});
	}

	private async loadAllSheets(body: HTMLElement, handle: WorksheetHandle) {
		// Re-read the file so newly-added sheets show up.
		const arrayBuffer = await this.app.vault.readBinary(handle.file);
		const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
		handle.workbook = workbook;

		// Build the new ranges list from the live workbook.
		const newRanges: SpreadsheetRange[] = [];
		for (const sheetName of workbook.SheetNames) {
			const worksheet = workbook.Sheets[sheetName];
			if (!worksheet || !worksheet['!ref']) continue; // skip empty / hidden
			const ref = XLSX.utils.decode_range(worksheet['!ref']);
			if (ref.e.r < ref.s.r || ref.e.c < ref.s.c) continue; // skip empty ranges
			newRanges.push({
				sheet: sheetName,
				startCell: XLSX.utils.encode_cell({ r: ref.s.r, c: ref.s.c }),
				endCell: XLSX.utils.encode_cell({ r: ref.e.r, c: ref.e.c }),
			});
		}

		if (newRanges.length === 0) {
			// Nothing to persist — fall back to a visible error in the block so
			// the user gets feedback even though the source didn't change.
			body.empty();
			body.createEl('div', {
				text: 'No sheets with data found in this workbook.',
				cls: 'spreadsheet-error',
			});
			return;
		}

		// Rewrite the code block source with the new ranges, keeping the
		// original options (`mode`, `formatting`) intact.
		const rangesPart = newRanges.map((r) => this.formatRange(r)).join(', ');
		const optionsPart = this.formatOptions(handle.spec.options);
		const newInnerArg = optionsPart ? `${rangesPart}; ${optionsPart}` : rangesPart;
		const newInnerText = `${handle.filename}(${newInnerArg})`;

		// Optimistic UI: render the new sub-blocks immediately, then persist
		// the rewrite. Obsidian's file watcher will trigger a re-render of
		// the block a tick later with the same content — harmless double-draw.
		body.empty();
		for (const range of newRanges) {
			this.renderSubBlock(body, handle, range, handle.spec.options);
		}

		try {
			await this.persistSpreadsheetSource(handle.el, handle.ctx, newInnerText);
		} catch (err) {
			const msg = err instanceof Error ? err.message : String(err);
			new Notice(`Could not update note source: ${msg}`);
		}
	}

	private formatRange(r: SpreadsheetRange): string {
		return r.sheet ? `${r.sheet}!${r.startCell}:${r.endCell}` : `${r.startCell}:${r.endCell}`;
	}

	private formatOptions(opts: RenderOptions): string {
		const parts: string[] = [];
		if (opts.mode !== 'stacked') parts.push(`mode=${opts.mode}`);
		if (!opts.formatting) parts.push('formatting=off');
		return parts.join(', ');
	}

	private async persistSpreadsheetSource(
		el: HTMLElement,
		ctx: MarkdownPostProcessorContext,
		newInnerText: string,
	): Promise<void> {
		// For a code block, the section's lineStart is the opening fence and
		// lineEnd is the closing fence. Splice the file content around that
		// range with our rewritten block.
		const section = ctx.getSectionInfo(el);
		if (!section) throw new Error('Code block section not found');

		const file = this.app.vault.getAbstractFileByPath(ctx.sourcePath);
		if (!(file instanceof TFile)) throw new Error(`Source file not found: ${ctx.sourcePath}`);

		const content = await this.app.vault.read(file);
		const lines = content.split('\n');
		const newLines = [
			...lines.slice(0, section.lineStart),
			'```spreadsheet',
			newInnerText,
			'```',
			...lines.slice(section.lineEnd + 1),
		];
		await this.app.vault.modify(file, newLines.join('\n'));
	}

	private renderSubBlock(
		parent: HTMLElement,
		handle: WorksheetHandle,
		range: SpreadsheetRange,
		options: RenderOptions,
	) {
		const sub = parent.createEl('div', { cls: 'spreadsheet-subblock' });

		// Per-range toolbar — sheet + range label.
		const subBar = sub.createEl('div', { cls: 'spreadsheet-subtoolbar' });
		const sheetName = range.sheet || handle.workbook.SheetNames[0];
		subBar.createEl('span', {
			cls: 'spreadsheet-label',
			text: `${sheetName} · ${range.startCell}:${range.endCell}`,
		});

		// Scroll wrapper.
		const scroll = sub.createEl('div', { cls: 'spreadsheet-scroll' });

		// Footer with Refresh + Use full range.
		const footer = sub.createEl('div', { cls: 'spreadsheet-footer' });
		this.renderFooter(footer, sub, handle, range, options);

		// Table itself.
		const table = scroll.createEl('table', { cls: 'spreadsheet-table' });
		const worksheet = handle.workbook.Sheets[sheetName];
		if (!worksheet) {
			sub.createEl('div', { text: `Sheet not found: ${sheetName}`, cls: 'spreadsheet-error' });
			return;
		}
		this.renderTable(table, worksheet, range, options.formatting);
	}

	private renderTabbedBlock(
		parent: HTMLElement,
		handle: WorksheetHandle,
		ranges: SpreadsheetRange[],
		options: RenderOptions,
	) {
		const tabBar = parent.createEl('div', { cls: 'spreadsheet-tabbar' });
		const subHosts: HTMLElement[] = [];
		const tabs: HTMLElement[] = [];

		ranges.forEach((range, idx) => {
			const sheetName = range.sheet || handle.workbook.SheetNames[0];
			const tabLabel = `${sheetName} · ${range.startCell}:${range.endCell}`;

			const tab = tabBar.createEl('button', {
				cls: 'spreadsheet-tab',
				text: tabLabel,
				attr: { type: 'button' },
			});
			if (idx === 0) tab.addClass('spreadsheet-tab-active');
			tabs.push(tab);

			const sub = parent.createEl('div', { cls: 'spreadsheet-subblock spreadsheet-tab-panel' });
			if (idx > 0) sub.style.display = 'none';
			subHosts.push(sub);

			tab.addEventListener('click', () => {
				subHosts.forEach((s, i) => {
					s.style.display = i === idx ? '' : 'none';
					tabs[i].removeClass('spreadsheet-tab-active');
				});
				tab.addClass('spreadsheet-tab-active');
			});

			// Sub-toolbar
			const subBar = sub.createEl('div', { cls: 'spreadsheet-subtoolbar' });
			subBar.createEl('span', { cls: 'spreadsheet-label', text: tabLabel });

			// Scroll wrapper
			const scroll = sub.createEl('div', { cls: 'spreadsheet-scroll' });

			// Footer
			const footer = sub.createEl('div', { cls: 'spreadsheet-footer' });
			this.renderFooter(footer, sub, handle, range, options);

			// Table
			const table = scroll.createEl('table', { cls: 'spreadsheet-table' });
			const worksheet = handle.workbook.Sheets[sheetName];
			if (!worksheet) {
				sub.createEl('div', { text: `Sheet not found: ${sheetName}`, cls: 'spreadsheet-error' });
				return;
			}
			this.renderTable(table, worksheet, range, options.formatting);
		});
	}

	private renderFooter(
		footer: HTMLElement,
		sub: HTMLElement,
		handle: WorksheetHandle,
		range: SpreadsheetRange,
		options: RenderOptions,
	) {
		footer.createEl('span', {
			cls: 'spreadsheet-label',
			text: `Range ${range.startCell}:${range.endCell}`,
		});

		const refreshBtn = footer.createEl('a', { text: 'Refresh', href: '#' });
		refreshBtn.addEventListener('click', async (evt) => {
			evt.preventDefault();
			await this.refreshSubBlock(sub, handle, range, options);
		});

		const fullBtn = footer.createEl('a', { text: 'Use full range', href: '#' });
		fullBtn.addEventListener('click', async (evt) => {
			evt.preventDefault();
			const sheetName = range.sheet || handle.workbook.SheetNames[0];
			const worksheet = handle.workbook.Sheets[sheetName];
			if (!worksheet || !worksheet['!ref']) return;
			const ref = XLSX.utils.decode_range(worksheet['!ref']);
			const fullRange: SpreadsheetRange = {
				sheet: range.sheet,
				startCell: XLSX.utils.encode_cell({ r: ref.s.r, c: ref.s.c }),
				endCell: XLSX.utils.encode_cell({ r: ref.e.r, c: ref.e.c }),
			};
			await this.refreshSubBlock(sub, handle, fullRange, options);
		});
	}

	private async refreshSubBlock(
		sub: HTMLElement,
		handle: WorksheetHandle,
		range: SpreadsheetRange,
		options: RenderOptions,
	) {
		// Re-read the file (may have changed on disk) and replace just the table.
		const arrayBuffer = await this.app.vault.readBinary(handle.file);
		const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
		handle.workbook = workbook;

		const sheetName = range.sheet || workbook.SheetNames[0];
		const worksheet = workbook.Sheets[sheetName];
		if (!worksheet) {
			sub.createEl('div', { text: `Sheet not found: ${sheetName}`, cls: 'spreadsheet-error' });
			return;
		}

		// Update the sub-toolbar label.
		const subBar = sub.querySelector('.spreadsheet-subtoolbar .spreadsheet-label');
		if (subBar) subBar.textContent = `${sheetName} · ${range.startCell}:${range.endCell}`;

		// Update the footer label.
		const footerLabel = sub.querySelector('.spreadsheet-footer .spreadsheet-label');
		if (footerLabel) footerLabel.textContent = `Range ${range.startCell}:${range.endCell}`;

		// Replace the table.
		const oldScroll = sub.querySelector('.spreadsheet-scroll');
		if (oldScroll) oldScroll.empty();
		const scroll = oldScroll || sub.createEl('div', { cls: 'spreadsheet-scroll' });
		const table = scroll.createEl('table', { cls: 'spreadsheet-table' });
		this.renderTable(table, worksheet, range, options.formatting);
	}

	private renderTable(
		table: HTMLElement,
		worksheet: XLSX.WorkSheet,
		range: SpreadsheetRange,
		formattingOn: boolean,
	) {
		const rangeInfo = XLSX.utils.decode_range(`${range.startCell}:${range.endCell}`);
		const headerRow = rangeInfo.s.r;

		for (let r = rangeInfo.s.r; r <= rangeInfo.e.r; r++) {
			const tr = table.createEl('tr');
			const isHeader = r === headerRow;
			for (let c = rangeInfo.s.c; c <= rangeInfo.e.c; c++) {
				const cellRef = XLSX.utils.encode_cell({ r, c });
				const cell = worksheet[cellRef];
				const td = tr.createEl(isHeader ? 'th' : 'td');
				if (cell?.v != null) {
					td.textContent = String(cell.v);
				}
				if (formattingOn && cell?.s) {
					this.applyCellStyle(td, cell.s);
				}
			}
		}
	}

	private applyCellStyle(td: HTMLElement, style: CellStyle) {
		const fillRgb = style?.fill?.fgColor?.rgb;
		if (fillRgb && typeof fillRgb === 'string') {
			// SheetJS may give 8-char ARGB; CSS wants 6-char RGB with a leading '#'.
			td.style.backgroundColor = '#' + fillRgb.slice(-6);
		}
		const fontRgb = style?.font?.color?.rgb;
		if (fontRgb && typeof fontRgb === 'string') {
			td.style.color = '#' + fontRgb.slice(-6);
		}
		if (style?.font?.bold) td.style.fontWeight = 'bold';
		if (style?.font?.italic) td.style.fontStyle = 'italic';
	}

	// ----------------------------------------------------------------------
	// Slash command / palette: insert a ```spreadsheet block at the cursor.
	// ----------------------------------------------------------------------

	private async openXlsxFilePicker(editor: Editor) {
		const files = this.app.vault
			.getFiles()
			.filter(f => /\.(xlsx|csv)$/i.test(f.path))
			.sort((a, b) => a.basename.localeCompare(b.basename));
		new XlsxFilePickerModal(this.app, files, (value) => this.handlePickedFile(editor, value)).open();
	}

	async handlePickedFile(editor: Editor, value: FileOrNew) {
		let file: TFile | null = null;
		if (value === CREATE_NEW) {
			file = await this.promptCreateXlsx();
		} else {
			file = value;
		}
		if (file) this.insertSpreadsheetBlock(editor, file);
	}

	insertSpreadsheetBlock(editor: Editor, file: TFile) {
		// If invoked from the slash menu, the editor still has the `/xlsx…`
		// trigger text in front of the cursor — strip it.
		const cursor = editor.getCursor();
		const before = editor.getLine(cursor.line).substring(0, cursor.ch);
		const triggerMatch = before.match(/\/[a-zA-Z][\w]*$/);
		const replaceStart = triggerMatch
			? { line: cursor.line, ch: cursor.ch - triggerMatch[0].length }
			: cursor;

		const blockText = '```spreadsheet\n' + file.name + '(A1:D10)\n```\n';
		editor.replaceRange(blockText, replaceStart, cursor);

		// Drop the cursor on the inner line, right after the opening paren,
		// so the user can immediately edit the range or sheet ref.
		editor.setCursor({
			line: replaceStart.line + 1,
			ch: file.name.length + 1,
		});
	}

	private async promptCreateXlsx(): Promise<TFile | null> {
		const defaultName = `New Sheet ${new Date().toISOString().slice(0, 10)}.xlsx`;
		const name = await new Promise<string | null>((resolve) => {
			const modal = new CreateXlsxModal(this.app, defaultName, resolve);
			modal.open();
		});
		if (!name) return null;

		// Drop the new file next to the note the user is editing — keeps the
		// path-relative block reference short and obvious.
		const activeFile = this.app.workspace.getActiveFile();
		const dir = activeFile?.parent?.path ?? '';
		const path = dir && dir !== '/' ? `${dir}/${name}` : name;

		if (this.app.vault.getAbstractFileByPath(path)) {
			new Notice(`File already exists: ${name}`);
			return null;
		}

		// Minimal valid xlsx: one empty sheet named "Sheet1".
		const wb = XLSX.utils.book_new();
		const ws = XLSX.utils.aoa_to_sheet([[]]);
		XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
		const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

		try {
			const file = await this.app.vault.createBinary(path, buf);
			new Notice(`Created ${name}`);
			return file;
		} catch (err) {
			const msg = err instanceof Error ? err.message : String(err);
			new Notice(`Could not create ${name}: ${msg}`);
			return null;
		}
	}
}

// ---------------------------------------------------------------------------
// Slash menu: appears while the user is typing `/xlsx…` in any note.
// ---------------------------------------------------------------------------

class XlsxFileSuggest extends EditorSuggest<FileOrNew> {
	plugin: SpreadsheetSyncPlugin;

	constructor(app: App, plugin: SpreadsheetSyncPlugin) {
		super(app);
		this.plugin = plugin;
	}

	onTrigger(cursor: EditorPosition, editor: Editor, _file: TFile | null): EditorSuggestTriggerInfo | null {
		// Trigger on `/x…` so we don't pop up for every stray slash. The regex
		// also lets the user keep typing past `/xlsx` (e.g. `/xlsxbudget`) to
		// narrow the file list — `query` is the part after the `/`.
		const line = editor.getLine(cursor.line);
		const before = line.substring(0, cursor.ch);
		const m = before.match(/\/([a-zA-Z][\w]*)$/);
		if (!m || !m[1].toLowerCase().startsWith('xlsx')) return null;
		return {
			start: { line: cursor.line, ch: cursor.ch - m[0].length },
			end: cursor,
			query: m[1],
		};
	}

	async getSuggestions(_ctx: EditorSuggestContext): Promise<FileOrNew[]> {
		const q = _ctx.query.substring(4).toLowerCase(); // strip the leading 'xlsx'
		const files = this.app.vault
			.getFiles()
			.filter(f => /\.(xlsx|csv)$/i.test(f.path) && f.basename.toLowerCase().includes(q))
			.sort((a, b) => a.basename.localeCompare(b.basename));
		return [CREATE_NEW, ...files];
	}

	renderSuggestion(value: FileOrNew, el: HTMLElement) {
		el.empty();
		if (value === CREATE_NEW) {
			el.addClass('spreadsheet-suggest-create');
			el.createEl('div', { text: '+ Create new xlsx file…' });
			el.createEl('small', { text: 'Empty workbook, one Sheet1' });
		} else {
			el.createEl('div', { text: value.basename, cls: 'spreadsheet-suggest-name' });
			el.createEl('small', { text: value.path, cls: 'spreadsheet-suggest-path' });
		}
	}

	async selectSuggestion(value: FileOrNew, _evt: MouseEvent | KeyboardEvent) {
		const editor = this.app.workspace.activeEditor?.editor;
		if (!editor) return;
		await this.plugin.handlePickedFile(editor, value);
	}
}

// ---------------------------------------------------------------------------
// Command palette picker: a FuzzySuggestModal with a "Create new" button.
// ---------------------------------------------------------------------------

class XlsxFilePickerModal extends FuzzySuggestModal<TFile> {
	files: TFile[];
	onSelect: (value: FileOrNew) => void | Promise<void>;

	constructor(app: App, files: TFile[], onSelect: (value: FileOrNew) => void | Promise<void>) {
		super(app);
		this.files = files;
		this.onSelect = onSelect;
		this.setPlaceholder('Search xlsx/csv files…');
	}

	getItems(): TFile[] {
		return this.files;
	}

	getItemText(item: TFile): string {
		return item.basename;
	}

	onOpen() {
		super.onOpen();
		// FuzzySuggestModal strips items that score 0, so a "+ Create new"
		// row inside `getItems()` would vanish the moment the user types
		// anything. Inject a button above the results instead.
		const prompt = this.modalEl.querySelector('.prompt-input');
		if (!prompt || !prompt.parentElement) return;
		const createBtn = prompt.parentElement.createEl('button', {
			text: '+ Create new xlsx file…',
			cls: 'spreadsheet-create-new-btn mod-cta',
		});
		createBtn.style.marginBottom = '8px';
		createBtn.style.width = '100%';
		createBtn.addEventListener('click', () => {
			this.close();
			void this.onSelect(CREATE_NEW);
		});
	}

	onChooseItem(item: TFile, _evt: MouseEvent | KeyboardEvent) {
		void this.onSelect(item);
	}
}

// ---------------------------------------------------------------------------
// Tiny modal: name input + Create / Cancel. Resolves to the typed name or
// null if dismissed.
// ---------------------------------------------------------------------------

class CreateXlsxModal extends Modal {
	defaultName: string;
	resolveFn: ((value: string | null) => void) | null;

	constructor(app: App, defaultName: string, resolveFn: (value: string | null) => void) {
		super(app);
		this.defaultName = defaultName;
		this.resolveFn = resolveFn;
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h3', { text: 'Create new xlsx file' });

		const input = contentEl.createEl('input', {
			type: 'text',
			value: this.defaultName,
			cls: 'spreadsheet-create-input',
		});
		input.style.width = '100%';
		input.focus();
		input.select();

		const row = contentEl.createDiv({ cls: 'spreadsheet-create-row' });
		row.style.marginTop = '12px';
		row.style.display = 'flex';
		row.style.gap = '8px';
		row.style.justifyContent = 'flex-end';
		const createBtn = row.createEl('button', { text: 'Create', cls: 'mod-cta' });
		const cancelBtn = row.createEl('button', { text: 'Cancel' });

		const submit = () => {
			const name = input.value.trim();
			if (!name) return;
			const final = /\.xlsx$/i.test(name) ? name : `${name}.xlsx`;
			const r = this.resolveFn;
			this.resolveFn = null;
			this.close();
			r?.(final);
		};
		const cancel = () => {
			const r = this.resolveFn;
			this.resolveFn = null;
			this.close();
			r?.(null);
		};

		createBtn.addEventListener('click', submit);
		cancelBtn.addEventListener('click', cancel);
		input.addEventListener('keydown', (e) => {
			if (e.key === 'Enter') {
				e.preventDefault();
				submit();
			} else if (e.key === 'Escape') {
				e.preventDefault();
				cancel();
			}
		});
	}

	onClose() {
		this.contentEl.empty();
		// If the modal was closed without submit/cancel firing (e.g. user
		// clicked the backdrop), resolve as a cancel.
		if (this.resolveFn) {
			const r = this.resolveFn;
			this.resolveFn = null;
			r(null);
		}
	}
}