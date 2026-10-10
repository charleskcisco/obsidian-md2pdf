import { App, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile } from 'obsidian';
import { exec, spawn } from 'child_process';
import { promisify } from 'util';
import { pathToFileURL } from 'url';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';

const execAsync = promisify(exec);

type PdfEngine = 'typst' | 'libreoffice';
type ExportFormat = 'pdf' | 'docx';

interface MD2PDFSettings {
	pandocPath: string;
	libreOfficePath: string;
	typstPath: string;
	pdfEngine: PdfEngine;
	defaultRef: string;
	deleteIntermediateDocx: boolean;
	autoDetectBibliography: boolean;
	customPandocArgs: string;
	pdfOutputLocation: string;
	docxOutputLocation: string;
	offerToOpen: boolean;
}

const DEFAULT_SETTINGS: MD2PDFSettings = {
	pandocPath: '',
	libreOfficePath: '',
	typstPath: '',
	pdfEngine: 'typst',
	defaultRef: '',
	deleteIntermediateDocx: true,
	autoDetectBibliography: true,
	customPandocArgs: '',
	pdfOutputLocation: '',
	docxOutputLocation: '',
	offerToOpen: true
};

// Properties the "Insert frontmatter" command adds, in this order.
const FRONTMATTER_PROPS = ['title', 'subtitle', 'author', 'instructor', 'course', 'date', 'spacing', 'style'];

// Spacings the Typst templates can express. Any other spacing: names a
// reference .docx (dg.double, quiz, ...) that only LibreOffice can honour.
const TYPST_SPACINGS = ['single', 'double'];

// Which .typ template renders which style:. Anything unrecognised falls
// back to the default rather than failing the export.
const DEFAULT_TEMPLATE = 'folio.typ';
const TEMPLATES: Record<string, string> = {
	mla: 'mla.typ',
	chicago: 'turabian.typ',
	turabian: 'turabian.typ'
};

// Citation style implied by style:, used when the note names no csl: of its
// own. Without this pandoc falls back to chicago-author-date, so a Turabian
// note asking for footnotes quietly got "(Smith 2020)" instead.
const DEFAULT_CSLS: Record<string, string> = {
	chicago: 'chicago-notes-bibliography.csl',
	turabian: 'chicago-notes-bibliography.csl',
	mla: 'modern-language-association.csl'
};

const MONTH_NAMES = [
	'January', 'February', 'March', 'April', 'May', 'June',
	'July', 'August', 'September', 'October', 'November', 'December'
];

// Generous: these exist to catch a hang, not to bound honest work.
const COMMAND_TIMEOUT_MS = 10 * 60 * 1000;

class CommandError extends Error {
	constructor(public stage: string, public args: string[], public code: number | null,
		public stdout: string, public stderr: string) {
		super(`${stage} failed${code !== null ? ` (exit ${code})` : ''}: ${lastLine(stderr || stdout) || 'no output'}`);
	}
}

function lastLine(s: string): string {
	const lines = s.trim().split('\n');
	return lines[lines.length - 1].slice(0, 120);
}

function styleKey(yaml: Record<string, string>): string {
	return (yaml['style'] || '').trim().toLowerCase();
}

function isChicago(yaml: Record<string, string>): boolean {
	const s = styleKey(yaml);
	return s === 'chicago' || s === 'turabian';
}

export default class MD2PDFPlugin extends Plugin {
	settings!: MD2PDFSettings;

	async onload() {
		await this.loadSettings();

		// Id kept from the original single command so existing hotkeys still work.
		this.addCommand({
			id: 'convert-md-to-pdf',
			name: 'Export current note to PDF',
			editorCallback: async (editor: Editor, view: MarkdownView | import('obsidian').MarkdownFileInfo) => {
				if (view.file) {
					await this.exportNote(view.file, 'pdf');
				} else {
					new Notice('No file is currently open');
				}
			}
		});

		this.addCommand({
			id: 'convert-md-to-docx',
			name: 'Export current note to DOCX',
			editorCallback: async (editor: Editor, view: MarkdownView | import('obsidian').MarkdownFileInfo) => {
				if (view.file) {
					await this.exportNote(view.file, 'docx');
				} else {
					new Notice('No file is currently open');
				}
			}
		});

		this.addCommand({
			id: 'insert-frontmatter',
			name: 'Insert export frontmatter',
			editorCallback: (editor: Editor) => this.insertFrontmatter(editor)
		});

		this.addRibbonIcon('file-output', 'Export note to PDF', async () => {
			const activeFile = this.app.workspace.getActiveFile();
			if (activeFile && activeFile.extension === 'md') {
				await this.exportNote(activeFile, 'pdf');
			} else {
				new Notice('Please open a Markdown file first');
			}
		});

		this.addSettingTab(new MD2PDFSettingTab(this.app, this));

		if (!this.settings.pandocPath || !this.settings.libreOfficePath || !this.settings.typstPath) {
			await this.autoDetectPaths();
		}
	}

	insertFrontmatter(editor: Editor) {
		const text = editor.getValue();
		const m = text.match(/^---\n([\s\S]*?)\n---/);
		if (m) {
			const existing = new Set<string>();
			for (const line of m[1].split('\n')) {
				const idx = line.indexOf(':');
				if (idx > 0) existing.add(line.substring(0, idx).trim());
			}
			const missing = FRONTMATTER_PROPS.filter(p => !existing.has(p));
			if (missing.length === 0) {
				new Notice('All frontmatter properties already present.');
				return;
			}
			// Append just before the closing fence; replaceRange keeps the cursor
			// where it was in the body.
			const endOfBlock = editor.offsetToPos(4 + m[1].length);
			editor.replaceRange('\n' + missing.map(p => `${p}: `).join('\n'), endOfBlock);
		} else {
			const block = FRONTMATTER_PROPS.map(p => `${p}: `).join('\n');
			editor.replaceRange(`---\n${block}\n---\n`, { line: 0, ch: 0 });
		}
		new Notice('Frontmatter inserted.');
	}

	async autoDetectPaths() {
		const platform = os.platform();
		let pandocPaths: string[] = [];
		let libreofficePaths: string[] = [];
		let typstPaths: string[] = [];

		if (platform === 'win32') {
			pandocPaths = [
				path.join('C:', 'Program Files', 'Pandoc', 'pandoc.exe'),
				path.join('C:', 'Program Files (x86)', 'Pandoc', 'pandoc.exe'),
				path.join(os.homedir(), 'AppData', 'Local', 'Pandoc', 'pandoc.exe'),
				path.join(os.homedir(), 'scoop', 'shims', 'pandoc.exe'),
				path.join('C:', 'ProgramData', 'chocolatey', 'bin', 'pandoc.exe')
			];
			libreofficePaths = [
				path.join('C:', 'Program Files', 'LibreOffice', 'program', 'soffice.exe'),
				path.join('C:', 'Program Files (x86)', 'LibreOffice', 'program', 'soffice.exe')
			];
			typstPaths = [
				path.join(os.homedir(), 'scoop', 'shims', 'typst.exe'),
				path.join(os.homedir(), '.cargo', 'bin', 'typst.exe'),
				path.join(os.homedir(), 'AppData', 'Local', 'Microsoft', 'WinGet', 'Links', 'typst.exe'),
				path.join('C:', 'ProgramData', 'chocolatey', 'bin', 'typst.exe')
			];
			// Also check for versioned LibreOffice folders
			const programFiles = process.env['ProgramFiles'] || path.join('C:', 'Program Files');
			const programFilesX86 = process.env['ProgramFiles(x86)'] || path.join('C:', 'Program Files (x86)');
			for (const base of [programFiles, programFilesX86]) {
				try {
					const entries = fs.readdirSync(base);
					for (const entry of entries) {
						if (entry.startsWith('LibreOffice')) {
							const sofficePath = path.join(base, entry, 'program', 'soffice.exe');
							if (!libreofficePaths.includes(sofficePath)) {
								libreofficePaths.push(sofficePath);
							}
						}
					}
				} catch (e) {
					// Directory doesn't exist or can't be read
				}
			}
		} else if (platform === 'darwin') {
			pandocPaths = [
				'/usr/local/bin/pandoc',
				'/opt/homebrew/bin/pandoc',
				'/usr/bin/pandoc'
			];
			libreofficePaths = [
				'/Applications/LibreOffice.app/Contents/MacOS/soffice',
				'/usr/local/bin/soffice'
			];
			typstPaths = [
				'/opt/homebrew/bin/typst',
				'/usr/local/bin/typst',
				path.join(os.homedir(), '.cargo', 'bin', 'typst')
			];
		} else {
			// Linux
			pandocPaths = [
				'/usr/bin/pandoc',
				'/usr/local/bin/pandoc',
				'/snap/bin/pandoc'
			];
			libreofficePaths = [
				'/usr/bin/libreoffice',
				'/usr/bin/soffice',
				'/usr/local/bin/libreoffice',
				'/snap/bin/libreoffice'
			];
			typstPaths = [
				'/usr/bin/typst',
				'/usr/local/bin/typst',
				'/snap/bin/typst',
				path.join(os.homedir(), '.cargo', 'bin', 'typst')
			];
		}

		if (!this.settings.pandocPath) {
			this.settings.pandocPath = await this.findBinary(pandocPaths,
				platform === 'win32' ? 'where pandoc' : 'which pandoc');
		}
		if (!this.settings.libreOfficePath) {
			this.settings.libreOfficePath = await this.findBinary(libreofficePaths,
				platform === 'win32' ? 'where soffice' : 'which libreoffice || which soffice');
		}
		if (!this.settings.typstPath) {
			this.settings.typstPath = await this.findBinary(typstPaths,
				platform === 'win32' ? 'where typst' : 'which typst');
		}

		await this.saveSettings();
	}

	async findBinary(candidates: string[], whichCmd: string): Promise<string> {
		for (const p of candidates) {
			if (fs.existsSync(p)) return p;
		}
		try {
			const { stdout } = await execAsync(whichCmd);
			const foundPath = stdout.trim().split('\n')[0].trim();
			if (foundPath && fs.existsSync(foundPath)) return foundPath;
		} catch (e) {
			// Command failed, path not found
		}
		return '';
	}

	async loadSettings() {
		const data = await this.loadData();
		this.settings = Object.assign({}, DEFAULT_SETTINGS, data);
		// An install that saved settings before the engine choice existed has
		// been exporting through LibreOffice all along. Leave it there rather
		// than switch a working setup's output on update; only fresh installs
		// start on Typst.
		if (data && !data.pdfEngine) {
			this.settings.pdfEngine = 'libreoffice';
		}
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	getVaultPath(): string {
		return (this.app.vault.adapter as any).basePath;
	}

	// Get plugin folder path
	getPluginFolder(): string {
		const vaultPath = this.getVaultPath();
		// manifest.dir is the folder Obsidian actually loaded us from, vault-relative.
		// The folder name does not have to match manifest.id — an install copied from
		// an older version still sits in a folder named 'md2pdf', and a zip from GitHub
		// unpacks as 'obsidian-md2pdf-main'. Deriving the path from the id looks in a
		// folder that does not exist and the reference docs come back empty.
		const pluginDir = this.manifest.dir
			|| path.join(this.app.vault.configDir, 'plugins', this.manifest.id);
		return path.join(vaultPath, pluginDir);
	}

	// Get list of available reference.docx files in plugin folder
	getAvailableRefDocs(): string[] {
		const pluginFolder = this.getPluginFolder();
		try {
			if (fs.existsSync(pluginFolder)) {
				const files = fs.readdirSync(pluginFolder);
				return files.filter(f => f.endsWith('.docx')).map(f => f.replace('.docx', ''));
			}
		} catch (e) {
			console.warn('Could not read plugin folder:', e);
		}
		return [];
	}

	// Parse YAML frontmatter from content. Deliberately splits on the first
	// colon rather than parsing real YAML: academic titles are full of
	// unquoted colons, which strict YAML rejects.
	parseYAML(content: string): Record<string, string> {
		const yaml: Record<string, string> = {};
		const match = content.match(/^---\r?\n([\s\S]*?)\r?\n---/);
		if (match) {
			const lines = match[1].split(/\r?\n/);
			for (const line of lines) {
				const colonIndex = line.indexOf(':');
				if (colonIndex > 0) {
					const key = line.substring(0, colonIndex).trim();
					let value = line.substring(colonIndex + 1).trim();
					if (value.length >= 2 && value[0] === value[value.length - 1] && (value[0] === '"' || value[0] === "'")) {
						value = value.slice(1, -1);
					}
					yaml[key] = value;
				}
			}
		}
		return yaml;
	}

	resolveReferenceDoc(yaml: Record<string, string>): string | null {
		const pluginFolder = this.getPluginFolder();
		if (yaml['spacing']) {
			const refPath = path.join(pluginFolder, yaml['spacing'] + '.docx');
			if (fs.existsSync(refPath)) return refPath;
			new Notice(`Reference doc not found: ${refPath}`);
			return null;
		}
		if (this.settings.defaultRef) {
			const defaultRefPath = path.join(pluginFolder, this.settings.defaultRef + '.docx');
			if (fs.existsSync(defaultRefPath)) return defaultRefPath;
		}
		// Final fallback to any .docx in plugin folder. Also covers a defaultRef that
		// points at a file this install does not have — settings sync across machines,
		// the .docx files do not always follow.
		const availableDocs = this.getAvailableRefDocs();
		if (availableDocs.includes('double')) return path.join(pluginFolder, 'double.docx');
		if (availableDocs.length > 0) return path.join(pluginFolder, availableDocs[0] + '.docx');

		new Notice(`No .docx reference document found in ${pluginFolder}`);
		console.error('md2pdf: no reference .docx in plugin folder', {
			pluginFolder,
			manifestDir: this.manifest.dir,
			manifestId: this.manifest.id,
			folderExists: fs.existsSync(pluginFolder)
		});
		return null;
	}

	templatePath(yaml: Record<string, string>): string {
		return path.join(this.getPluginFolder(), 'templates', TEMPLATES[styleKey(yaml)] || DEFAULT_TEMPLATE);
	}

	// Which engine renders PDF for this note. A spacing: the Typst templates
	// cannot express names a reference .docx, so only LibreOffice can honour
	// it; otherwise the setting decides, and Typst is only used when both the
	// binary and the template are actually there.
	pdfEngineFor(yaml: Record<string, string>): PdfEngine {
		const spacing = yaml['spacing'];
		if (spacing && !TYPST_SPACINGS.includes(spacing)) return 'libreoffice';
		if (this.settings.pdfEngine === 'libreoffice') return 'libreoffice';
		const typstOk = !!this.settings.typstPath && fs.existsSync(this.settings.typstPath)
			&& fs.existsSync(this.templatePath(yaml));
		return typstOk ? 'typst' : 'libreoffice';
	}

	resolveOutputDir(setting: string, fallback: string): string {
		if (!setting) return fallback;
		const dir = path.isAbsolute(setting) ? setting : path.join(this.getVaultPath(), setting);
		fs.mkdirSync(dir, { recursive: true });
		return dir;
	}

	// Locate the .bib named by bibliography:. Tries the literal value
	// (absolute, then vault-relative), then the vault's own .bib — so a bare
	// `bibliography: true`, or an absolute path written on another machine,
	// still finds it.
	resolveBibPath(yaml: Record<string, string>): string | null {
		const vaultPath = this.getVaultPath();
		const val = (yaml['bibliography'] || '').trim();
		if (val) {
			for (const cand of [expandHome(val), path.join(vaultPath, val)]) {
				if (isFile(cand)) return cand;
			}
		}
		return findBibFile(vaultPath);
	}

	// Locate the .csl for this note. An explicit csl: always wins; only when
	// the note names none (or names one that is not on this machine) does the
	// style: default apply.
	resolveCslPath(yaml: Record<string, string>): string | null {
		const vaultPath = this.getVaultPath();
		const val = (yaml['csl'] || '').trim();
		if (val) {
			const name = path.basename(val);
			for (const cand of [expandHome(val), path.join(vaultPath, val),
				path.join(vaultPath, 'sources', name), path.join(vaultPath, name)]) {
				if (isFile(cand)) return cand;
			}
		}
		const def = DEFAULT_CSLS[styleKey(yaml)];
		if (!def) return null;
		const p = path.join(this.getPluginFolder(), 'csl', def);
		return isFile(p) ? p : null;
	}

	async exportNote(file: TFile, format: ExportFormat): Promise<void> {
		const content = await this.app.vault.read(file);
		const yaml = this.parseYAML(content);
		const engine: PdfEngine = format === 'pdf' ? this.pdfEngineFor(yaml) : 'libreoffice';

		// Validate the tools this export actually needs
		if (!this.settings.pandocPath || !fs.existsSync(this.settings.pandocPath)) {
			new Notice(this.settings.pandocPath
				? `Pandoc not found at: ${this.settings.pandocPath}`
				: 'Pandoc path not configured. Please set it in plugin settings.');
			return;
		}
		if (format === 'pdf' && engine === 'libreoffice'
			&& (!this.settings.libreOfficePath || !fs.existsSync(this.settings.libreOfficePath))) {
			new Notice(this.settings.libreOfficePath
				? `LibreOffice not found at: ${this.settings.libreOfficePath}`
				: 'LibreOffice path not configured. Install Typst or LibreOffice, or set the path in plugin settings.');
			return;
		}

		let referenceDoc: string | null = null;
		if (engine === 'libreoffice') {
			referenceDoc = this.resolveReferenceDoc(yaml);
			if (!referenceDoc) return;
		}

		// A note that asks for a bibliography and does not get one must not
		// export anyway: citations would degrade to literal [@key] (docx) or
		// abort the compile with an unhelpful message (Typst).
		const wantsBib = this.settings.autoDetectBibliography && 'bibliography' in yaml;
		const bibSrc = wantsBib ? this.resolveBibPath(yaml) : null;
		if (wantsBib && !bibSrc) {
			new Notice(`bibliography: not found — ${yaml['bibliography']}`);
			return;
		}
		const cslSrc = bibSrc ? this.resolveCslPath(yaml) : null;

		const vaultPath = this.getVaultPath();
		const inputPath = path.join(vaultPath, file.path);
		const noteDir = path.dirname(inputPath);
		const baseName = path.basename(file.path, '.md');

		const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'md2pdf-'));
		const progress = new Notice('', 0);
		const step = (msg: string) => progress.setMessage(`Exporting ${baseName}… ${msg}`);

		try {
			const pdfOutputDir = this.resolveOutputDir(this.settings.pdfOutputLocation, noteDir);
			const docxOutputDir = this.resolveOutputDir(this.settings.docxOutputLocation, noteDir);
			const pdfPath = path.join(pdfOutputDir, `${baseName}.pdf`);
			const keepDocx = format === 'docx' || !this.settings.deleteIntermediateDocx;
			const docxPath = path.join(keepDocx ? docxOutputDir : tmpDir, `${baseName}.docx`);

			// Images and other relative links still resolve from the note's own
			// folder, and from the vault root, now the source sits in a temp dir.
			const commonArgs = [`--resource-path=${[noteDir, vaultPath].join(path.delimiter)}`];
			if (bibSrc) {
				// Passed explicitly because pandoc resolves the frontmatter's own
				// values against the working directory, not the vault.
				const bibUse = narrowBib(bibSrc, content, tmpDir);
				commonArgs.push('--citeproc', `--bibliography=${bibUse}`);
				if (cslSrc) commonArgs.push(`--csl=${cslSrc}`);
			}
			const customArgs = this.settings.customPandocArgs.split(/\s+/).filter(arg => arg.length > 0);

			if (engine === 'typst') {
				// pandoc parses the markdown only; the template does all the
				// formatting from metadata we hand it ourselves. The frontmatter
				// is stripped so pandoc's strict YAML parser never sees it.
				const bodyMd = path.join(tmpDir, 'body.md');
				fs.writeFileSync(bodyMd, wrapBibliography(stripFrontmatter(content)), 'utf8');

				step(bibSrc ? '(1/2) Running pandoc with citations' : '(1/2) Running pandoc');
				const pandocArgs = [bodyMd, ...commonArgs];
				// Without a bibliography, keep @keys literal; otherwise Typst
				// aborts on an unresolvable label.
				if (!bibSrc) pandocArgs.push('--from', 'markdown-citations');
				// typst-citations: emit citeproc's rendered text rather than
				// native #cite() calls, which need a #bibliography() the
				// templates deliberately lack.
				pandocArgs.push('--extract-media=media', ...customArgs,
					'--to', 'typst-citations', '-o', 'body.typ');
				await this.runCommand('pandoc', this.settings.pandocPath, pandocArgs, tmpDir);

				fs.copyFileSync(this.templatePath(yaml), path.join(tmpDir, 'folio.typ'));
				fs.writeFileSync(path.join(tmpDir, 'main.typ'), typstWrapper(yaml, 'body.typ'), 'utf8');

				step('(2/2) Rendering PDF');
				const typstArgs = ['compile', '--root', tmpDir];
				const fontsDir = path.join(this.getPluginFolder(), 'fonts');
				if (fs.existsSync(fontsDir)) typstArgs.push('--font-path', fontsDir);
				typstArgs.push(path.join(tmpDir, 'main.typ'), pdfPath);
				await this.runCommand('typst', this.settings.typstPath, typstArgs, tmpDir);
				if (!fs.existsSync(pdfPath)) throw new Error('typst ran but produced no PDF');

				progress.hide();
				this.exported(pdfPath);
				return;
			}

			// The frontmatter is stripped here too, for the same reason as the
			// Typst path: the Lua filters take their metadata from our own
			// parse, and an unstyled note gets its title block via -M instead.
			const sourceMd = path.join(tmpDir, 'source.md');
			fs.writeFileSync(sourceMd, wrapBibliography(stripFrontmatter(content)), 'utf8');
			const metaArgs: string[] = [];
			if (styleKey(yaml) !== 'mla' && !isChicago(yaml)) {
				for (const k of ['title', 'subtitle', 'author', 'date']) {
					if (yaml[k]) metaArgs.push(`--metadata=${k}:${yaml[k]}`);
				}
			}
			const luaPath = path.join(tmpDir, 'filter.lua');
			fs.writeFileSync(luaPath, this.generateLuaFilter(yaml), 'utf8');

			const total = format === 'pdf' ? 3 : 2;
			step(`(1/${total}) Running pandoc${bibSrc ? ' with citations' : ''}`);
			await this.runCommand('pandoc', this.settings.pandocPath, [
				sourceMd,
				'--standalone',
				'--reference-doc=' + referenceDoc,
				'--lua-filter=' + luaPath,
				...metaArgs,
				...commonArgs,
				...customArgs,
				'-o', docxPath
			], noteDir);
			if (!fs.existsSync(docxPath)) throw new Error('Pandoc failed to create DOCX file');

			step(`(2/${total}) Post-processing`);
			this.postprocessDocx(docxPath, yaml);

			if (format === 'docx') {
				progress.hide();
				this.exported(docxPath);
				return;
			}

			step('(3/3) Converting to PDF');
			// A private, writable profile: headless soffice commonly exits 0
			// having written nothing when its default profile is locked — for
			// instance because LibreOffice is already open.
			const loProfile = pathToFileURL(path.join(tmpDir, 'loprofile')).href;
			await this.runCommand('LibreOffice', this.settings.libreOfficePath, [
				`-env:UserInstallation=${loProfile}`,
				'--headless', '--norestore',
				'--convert-to', 'pdf',
				'--outdir', pdfOutputDir,
				docxPath
			], pdfOutputDir);
			if (!fs.existsSync(pdfPath)) throw new Error('LibreOffice ran but produced no PDF');

			progress.hide();
			this.exported(pdfPath);
		} catch (error) {
			progress.hide();
			console.error('MD2PDF export error:', error);
			let msg = `Export failed: ${(error as Error).message}`;
			if (error instanceof CommandError) {
				const log = this.logExportError(error, file.path, format);
				if (log) msg += ` (see ${log})`;
			}
			new Notice(msg, 15000);
		} finally {
			fs.rmSync(tmpDir, { recursive: true, force: true });
		}
	}

	exported(outPath: string) {
		new Notice(`Exported: ${path.basename(outPath)}`);
		if (this.settings.offerToOpen) {
			new AfterExportModal(this.app, outPath).open();
		}
	}

	// Notices are fleeting and truncated, so write the full command and output
	// somewhere durable. A dot-folder keeps it out of the file explorer.
	logExportError(err: CommandError, note: string, format: ExportFormat): string | null {
		try {
			const rel = path.join('.md2pdf', 'export-error.log');
			const logPath = path.join(this.getVaultPath(), rel);
			fs.mkdirSync(path.dirname(logPath), { recursive: true });
			fs.writeFileSync(logPath,
				`md2pdf export failed\n` +
				`stage:   ${err.stage}\n` +
				`format:  ${format}\n` +
				`note:    ${note}\n` +
				`command:\n  ${err.args.join(' ')}\n` +
				`returncode: ${err.code}\n` +
				`--- stdout ---\n${err.stdout}\n` +
				`--- stderr ---\n${err.stderr}\n`, 'utf8');
			return rel;
		} catch (e) {
			return null;
		}
	}

	// One pass over the docx: strip headers or footers by style, fill in the
	// {{LASTNAME}} placeholder, and re-size table columns to their contents.
	// style: mla -> keep headers, strip footers
	// style: chicago or blank -> keep footers, strip headers
	postprocessDocx(docxPath: string, yaml: Record<string, string>): void {
		const AdmZip = require('adm-zip');
		const isMla = styleKey(yaml) === 'mla';
		const lastName = authorLastname(yaml);

		try {
			const zip = new AdmZip(docxPath);
			for (const entry of zip.getEntries()) {
				const name: string = entry.entryName;
				const isHeader = /^word\/header\d*\.xml$/.test(name);
				const isFooter = /^word\/footer\d*\.xml$/.test(name);
				let content: string | null = null;

				if (name === 'word/document.xml') {
					content = autofitTables(entry.getData().toString('utf8'));
				} else if (isHeader && !isMla) {
					content = EMPTY_HEADER_XML;
				} else if (isFooter && isMla) {
					content = EMPTY_FOOTER_XML;
				} else if (isHeader || isFooter) {
					const text: string = entry.getData().toString('utf8');
					content = lastName
						? text.split('{{LASTNAME}}').join(lastName)
						: text.split('{{LASTNAME}} ').join('').split('{{LASTNAME}}').join('');
				}
				if (content !== null) zip.updateFile(name, Buffer.from(content, 'utf8'));
			}
			zip.writeZip(docxPath);
		} catch (e) {
			console.warn('Failed to post-process docx:', e);
			// Non-fatal - continue with conversion
		}
	}

	generateLuaFilter(yaml: Record<string, string>): string {
		const s = styleKey(yaml);
		if (s === 'mla') return this.generateHeaderFilter(yaml);
		if (isChicago(yaml)) return this.generateCoverPageFilter(yaml);
		return this.generateBasicFilter();
	}

	luaMetaVars(yaml: Record<string, string>): string {
		const keys = ['title', 'subtitle', 'author', 'course', 'instructor', 'date'];
		return `local M = {\n${keys.map(k => `  ${k} = ${luaStr(yaml[k] || '')},`).join('\n')}\n}

-- Capture from document metadata when not in the frontmatter we parsed,
-- then remove so pandoc does not also render them in its own styles.
function Meta(meta)
  for k, v in pairs(M) do
    if meta[k] and v == "" then
      M[k] = pandoc.utils.stringify(meta[k])
    end
    meta[k] = nil
  end
  return meta
end`;
	}

	luaBibHelpers(): string {
		return `local function escape_xml(s)
  s = s:gsub("&", "&amp;")
  s = s:gsub("<", "&lt;")
  s = s:gsub(">", "&gt;")
  return s
end

-- A plain double-spaced paragraph: before is spacing in twips, center a bool.
local function para_xml(text, before, center)
  return pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:before="%d" w:after="0" w:line="480" w:lineRule="auto"/>%s
  </w:pPr>
  <w:r>
    <w:t xml:space="preserve">%s</w:t>
  </w:r>
</w:p>]], before or 0, center and '\\n    <w:jc w:val="center"/>' or '', escape_xml(text)))
end

local function format_date(date_str, mla)
  local months = {
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  }
  local year, month, day = date_str:match("(%d+)-(%d+)-(%d+)")
  if year and month and day then
    local month_name = months[tonumber(month)]
    if month_name then
      if mla then
        return string.format("%d %s %s", tonumber(day), month_name, year)
      end
      return string.format("%s %d, %s", month_name, tonumber(day), year)
    end
  end
  return date_str
end

local function inlines_to_openxml(inlines)
  local runs = {}
  for _, inl in ipairs(inlines) do
    if inl.t == "Emph" then
      local txt = escape_xml(pandoc.utils.stringify(inl))
      table.insert(runs, string.format(
        '<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>', txt))
    elseif inl.t == "Strong" then
      local txt = escape_xml(pandoc.utils.stringify(inl))
      table.insert(runs, string.format(
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>', txt))
    elseif inl.t == "Str" then
      table.insert(runs, string.format(
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>', escape_xml(inl.text)))
    elseif inl.t == "Space" then
      table.insert(runs, '<w:r><w:t xml:space="preserve"> </w:t></w:r>')
    elseif inl.t == "SoftBreak" or inl.t == "LineBreak" then
      table.insert(runs, '<w:r><w:t xml:space="preserve"> </w:t></w:r>')
    elseif inl.t == "Link" then
      local txt = escape_xml(pandoc.utils.stringify(inl))
      table.insert(runs, string.format(
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>', txt))
    else
      local txt = escape_xml(pandoc.utils.stringify(inl))
      if txt ~= "" then
        table.insert(runs, string.format(
          '<w:r><w:t xml:space="preserve">%s</w:t></w:r>', txt))
      end
    end
  end
  return table.concat(runs)
end

local function bib_entry_block(block)
  local runs_xml = inlines_to_openxml(block.content)
  return pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:after="0" w:line="480" w:lineRule="auto"/>
    <w:ind w:left="720" w:hanging="720"/>
  </w:pPr>
  %s
</w:p>]], runs_xml))
end

local function is_bib_heading(block)
  if block.t ~= "Header" then return false end
  local text = pandoc.utils.stringify(block):lower()
  return text:match("bibliography") or text:match("references") or text:match("works cited")
end

local function is_refs_div(block)
  if block.t ~= "Div" then return false end
  if block.identifier == "bibentries" then return true end
  for _, cls in ipairs(block.classes) do
    if cls == "references" or cls == "csl-bib-body" then return true end
  end
  return false
end

-- Returns true if the block was consumed as part of a bibliography.
local bib_state = { in_bib = false }
local function handle_bib_block(block, new_blocks)
  if is_bib_heading(block) then
    bib_state.in_bib = true
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading%d"/>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t xml:space="preserve">%s</w:t>
  </w:r>
</w:p>]], block.level, escape_xml(pandoc.utils.stringify(block)))))
    return true
  elseif bib_state.in_bib and block.t == "Header" then
    bib_state.in_bib = false
    return false
  elseif bib_state.in_bib and block.t == "Para" then
    table.insert(new_blocks, bib_entry_block(block))
    return true
  elseif (bib_state.in_bib or is_refs_div(block)) and block.t == "Div" then
    for _, inner in ipairs(block.content) do
      if inner.t == "Div" then
        for _, b in ipairs(inner.content) do
          if b.t == "Para" then
            table.insert(new_blocks, bib_entry_block(b))
          end
        end
      elseif inner.t == "Para" then
        table.insert(new_blocks, bib_entry_block(inner))
      end
    end
    return true
  end
  return false
end`;
	}

	generateBasicFilter(): string {
		return `${this.luaBibHelpers()}

function Pandoc(doc)
  local new_blocks = {}
  for _, block in ipairs(doc.blocks) do
    if not handle_bib_block(block, new_blocks) then
      table.insert(new_blocks, block)
    end
  end
  doc.blocks = new_blocks
  return doc
end`;
	}

	// Turabian cover page: title centered ~1/3 down, author/course/instructor/
	// date in the bottom half, body from page two.
	generateCoverPageFilter(yaml: Record<string, string>): string {
		return `${this.luaBibHelpers()}
${this.luaMetaVars(yaml)}

function Pandoc(doc)
  local new_blocks = {}

  if M.title ~= "" then
    -- With a subtitle, the title line ends with a colon and the subtitle
    -- takes the line below. A title typed with its own colon keeps just one.
    local title_line = M.title
    if M.subtitle ~= "" and not M.title:match(":%s*$") then
      title_line = M.title .. ":"
    end
    table.insert(new_blocks, para_xml(title_line, 2400, true))
    if M.subtitle ~= "" then
      table.insert(new_blocks, para_xml(M.subtitle, 0, true))
    end
  end

  -- Roughly the bottom third of the page
  local gap = 4320
  local info = { M.author, M.course, M.instructor,
                 M.date ~= "" and format_date(M.date, false) or "" }
  for _, line in ipairs(info) do
    if line ~= "" then
      table.insert(new_blocks, para_xml(line, gap, true))
      gap = 0
    end
  end

  local page_break_inserted = false
  for _, block in ipairs(doc.blocks) do
    if is_bib_heading(block) then
      page_break_inserted = true
    end
    if not handle_bib_block(block, new_blocks) then
      if not page_break_inserted and (block.t == "Header" or
         (block.t == "Para" and #block.content > 0) or
         block.t == "CodeBlock" or block.t == "BulletList" or
         block.t == "OrderedList" or block.t == "Table" or
         block.t == "BlockQuote" or block.t == "RawBlock" or block.t == "Div") then
        table.insert(new_blocks, pandoc.RawBlock('openxml',
          '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'))
        page_break_inserted = true
      end
      table.insert(new_blocks, block)
    end
  end

  if not page_break_inserted then
    table.insert(new_blocks, pandoc.RawBlock('openxml',
      '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'))
  end

  doc.blocks = new_blocks
  return doc
end`;
	}

	// MLA: author, instructor, course, date left-aligned, then a centered
	// title. The running "LastName PageNumber" head comes from the reference doc.
	generateHeaderFilter(yaml: Record<string, string>): string {
		return `${this.luaBibHelpers()}
${this.luaMetaVars(yaml)}

function Pandoc(doc)
  local new_blocks = {}

  local info = { M.author, M.instructor, M.course,
                 M.date ~= "" and format_date(M.date, true) or "" }
  for _, line in ipairs(info) do
    if line ~= "" then
      table.insert(new_blocks, para_xml(line, 0, false))
    end
  end

  if M.title ~= "" then
    -- MLA keeps title and subtitle on one line, joined by a colon.
    local full_title = M.title
    if M.subtitle ~= "" then
      if M.title:match(":%s*$") then
        full_title = M.title:gsub("%s*$", "") .. " " .. M.subtitle
      else
        full_title = M.title .. ": " .. M.subtitle
      end
    end
    table.insert(new_blocks, para_xml(full_title, 0, true))
  end

  for _, block in ipairs(doc.blocks) do
    if not handle_bib_block(block, new_blocks) then
      table.insert(new_blocks, block)
    end
  end

  doc.blocks = new_blocks
  return doc
end`;
	}

	runCommand(stage: string, command: string, args: string[], cwd: string): Promise<void> {
		return new Promise((resolve, reject) => {
			const isWindows = os.platform() === 'win32';

			let proc;
			if (isWindows) {
				// On Windows, quote the command and use shell
				const quotedCommand = `"${command}"`;
				const quotedArgs = args.map(arg => arg.includes(' ') ? `"${arg}"` : arg);
				proc = spawn(quotedCommand, quotedArgs, { cwd, shell: true });
			} else {
				proc = spawn(command, args, { cwd });
			}

			let stderr = '';
			let stdout = '';
			const timer = setTimeout(() => proc.kill(), COMMAND_TIMEOUT_MS);

			proc.stdout.on('data', (data) => {
				stdout += data.toString();
			});

			proc.stderr.on('data', (data) => {
				stderr += data.toString();
			});

			proc.on('close', (code, signal) => {
				clearTimeout(timer);
				if (code === 0) {
					resolve();
				} else {
					if (signal) stderr += `\nkilled by ${signal}`;
					reject(new CommandError(stage, [command, ...args], code, stdout, stderr));
				}
			});

			proc.on('error', (err) => {
				clearTimeout(timer);
				reject(err);
			});
		});
	}

	onunload() {
		// Cleanup if needed
	}
}

// ── Helpers shared by both engines ─────────────────────────────────────

function expandHome(p: string): string {
	return p.startsWith('~') ? path.join(os.homedir(), p.slice(1)) : p;
}

function isFile(p: string): boolean {
	try {
		return fs.statSync(p).isFile();
	} catch (e) {
		return false;
	}
}

// A .bib in the vault: sources/ first, then anywhere.
function findBibFile(vaultPath: string): string | null {
	const sources = path.join(vaultPath, 'sources');
	try {
		const hit = fs.readdirSync(sources).sort().find(f => f.endsWith('.bib') && !f.startsWith('._'));
		if (hit) return path.join(sources, hit);
	} catch (e) {
		// No sources/ folder
	}
	const walk = (dir: string): string | null => {
		let entries: fs.Dirent[];
		try {
			entries = fs.readdirSync(dir, { withFileTypes: true });
		} catch (e) {
			return null;
		}
		entries.sort((a, b) => a.name.localeCompare(b.name));
		for (const e of entries) {
			if (e.isFile() && e.name.endsWith('.bib') && !e.name.startsWith('._')) return path.join(dir, e.name);
		}
		for (const e of entries) {
			if (e.isDirectory() && !e.name.startsWith('.') && e.name !== 'node_modules') {
				const found = walk(path.join(dir, e.name));
				if (found) return found;
			}
		}
		return null;
	};
	return walk(vaultPath);
}

function luaStr(s: string): string {
	return '"' + s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r?\n/g, '\\n') + '"';
}

function typstStr(s: string): string {
	return '"' + s.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
}

// Surname for the running head: explicit lastname: wins, else the last word
// of author:. Shared by both engines so they never disagree.
function authorLastname(yaml: Record<string, string>): string {
	if (yaml['lastname']) return yaml['lastname'];
	const parts = (yaml['author'] || '').split(/\s+/).filter(Boolean);
	return parts.length ? parts[parts.length - 1] : '';
}

// Turabian writes "January 1, 2026", MLA "1 January 2026". Anything that is
// not YYYY-MM-DD passes through untouched.
function formatExportDate(yaml: Record<string, string>): string {
	const raw = yaml['date'] || '';
	const m = raw.match(/(\d+)-(\d+)-(\d+)/);
	if (!m) return raw;
	const month = parseInt(m[2], 10);
	const day = parseInt(m[3], 10);
	if (month < 1 || month > 12) return raw;
	const name = MONTH_NAMES[month - 1];
	return styleKey(yaml) === 'mla' ? `${day} ${name} ${m[1]}` : `${name} ${day}, ${m[1]}`;
}

function stripFrontmatter(content: string): string {
	const m = content.match(/^---\r?\n[\s\S]*?\r?\n---[ \t]*\r?\n?/);
	return m ? content.slice(m[0].length) : content;
}

function typstWrapper(yaml: Record<string, string>, bodyName: string): string {
	const fields: [string, string][] = [
		['title', yaml['title'] || ''],
		['subtitle', yaml['subtitle'] || ''],
		['author', yaml['author'] || ''],
		['course', yaml['course'] || ''],
		['instructor', yaml['instructor'] || ''],
		['date', formatExportDate(yaml)],
		['style', yaml['style'] || ''],
		['spacing', yaml['spacing'] || 'double'],
		['lastname', authorLastname(yaml)]
	];
	const args = fields.map(([k, v]) => `  ${k}: ${typstStr(v)},\n`).join('');
	return `#import "folio.typ": conf\n#show: conf.with(\n${args})\n#include ${typstStr(bodyName)}\n`;
}

// Headings that begin a hand-typed reference list.
const BIB_HEADING_RE = /^(#{1,6})\s*(bibliography|works\s+cited|references)\s*$/im;

// Fence a hand-typed reference list so both engines can give it a hanging
// indent: the id becomes a <bibentries> label the Typst templates select,
// and custom-style names the reference docs' Bibliography paragraph style.
// Untouched if there is no such heading, the section is empty, or it
// already carries a fence.
function wrapBibliography(markdown: string): string {
	const m = BIB_HEADING_RE.exec(markdown);
	if (!m) return markdown;
	const start = m.index + m[0].length;
	const rest = markdown.slice(start);
	const nxt = /^#{1,6}\s/m.exec(rest);
	const end = start + (nxt ? nxt.index : rest.length);
	const body = markdown.slice(start, end);
	if (!body.trim() || body.includes(':::')) return markdown;
	return markdown.slice(0, start)
		+ '\n\n::: {#bibentries custom-style="Bibliography"}\n'
		+ body.replace(/^\n+|\n+$/g, '') + '\n:::\n' + markdown.slice(end);
}

// ── Bibliography narrowing ─────────────────────────────────────────────
//
// citeproc parses the whole bibliography before rendering anything, so an
// 800-entry library costs the same whether a note cites four sources or four
// hundred. Hand pandoc only what the note cites — and fall back to the
// original on any doubt, since silently dropping a source is far worse.

const CITEKEY_RE = /@([A-Za-z0-9_][A-Za-z0-9_:.#$%&+?<>~/-]*)/g;
const BIB_GLOBALS = ['string', 'preamble'];

function citedKeys(markdown: string): Set<string> {
	const keys = new Set<string>();
	for (const m of markdown.matchAll(CITEKEY_RE)) keys.add(m[1].replace(/[.,;:]+$/, ''));
	return keys;
}

interface BibBlock { kind: string; key: string | null; body: string; }

// Brace matching rather than line matching: an entry spans as many lines as
// it likes. Only the delimiter the entry opened with is counted, since prose
// in an abstract is full of unbalanced parentheses.
function bibBlocks(text: string): BibBlock[] {
	const blocks: BibBlock[] = [];
	let i = 0;
	const n = text.length;
	while (true) {
		const at = text.indexOf('@', i);
		if (at < 0) break;
		const m = /^@\s*([A-Za-z]+)\s*([{(])/.exec(text.slice(at, at + 64));
		if (!m) {
			i = at + 1;
			continue;
		}
		const kind = m[1].toLowerCase();
		const opener = m[2];
		const closer = opener === '{' ? '}' : ')';
		let depth = 0;
		let j = at + m[0].length - 1;
		while (j < n) {
			const c = text[j];
			if (c === '\\') {
				j += 2;
				continue;
			}
			if (c === opener) depth++;
			else if (c === closer) {
				depth--;
				if (depth === 0) {
					j++;
					break;
				}
			}
			j++;
		}
		const body = text.slice(at, j);
		let key: string | null = null;
		if (!BIB_GLOBALS.includes(kind)) {
			const km = /^@\s*[A-Za-z]+\s*[{(]\s*([^,\s}]+)/.exec(body);
			key = km ? km[1] : null;
		}
		blocks.push({ kind, key, body });
		i = j;
	}
	return blocks;
}

function filterBib(text: string, keys: Set<string>): string | null {
	const blocks = bibBlocks(text);
	if (!blocks.length) return null;
	const byKey = new Map<string, string>();
	for (const b of blocks) {
		if (!b.key) continue;
		if (!byKey.has(b.key)) byKey.set(b.key, b.body);
		if (!byKey.has(b.key.toLowerCase())) byKey.set(b.key.toLowerCase(), b.body);
	}

	// crossref and xdata name further entries, transitively.
	const wanted = new Set<string>();
	const queue = [...keys];
	while (queue.length) {
		const k = queue.pop()!;
		const body = byKey.get(k) ?? byKey.get(k.toLowerCase());
		if (body === undefined || wanted.has(k)) continue;
		wanted.add(k);
		for (const m of body.matchAll(/(?:crossref|xdata)\s*=\s*[{"]([^}"]+)/gi)) {
			queue.push(...m[1].split(',').map(r => r.trim()));
		}
	}
	if (!wanted.size) return null;

	const wantedLower = new Set([...wanted].map(w => w.toLowerCase()));
	const out = blocks
		.filter(b => BIB_GLOBALS.includes(b.kind) || (b.key && wantedLower.has(b.key.toLowerCase())))
		.map(b => b.body);
	const result = out.join('\n\n') + '\n';

	// Safety net: every cited key that exists in the original must survive.
	const kept = new Set(bibBlocks(result).filter(b => b.key).map(b => b.key!.toLowerCase()));
	for (const k of keys) {
		const present = byKey.has(k) || byKey.has(k.toLowerCase());
		if (present && !kept.has(k.toLowerCase())) return null;
	}
	return result;
}

// Fields no citation style renders, and the bulk of a Zotero export by volume.
// Dropping them roughly halves the time citeproc spends parsing.
const BIB_NOISE_RE = /([,{]\s*)(?:abstract|file|keywords)\s*=\s*/gi;

function bibValueEnd(text: string, i: number): number | null {
	const n = text.length;
	if (i >= n) return null;
	if (text[i] === '{') {
		let depth = 0;
		while (i < n) {
			const c = text[i];
			if (c === '\\') {
				i += 2;
				continue;
			}
			if (c === '{') depth++;
			else if (c === '}') {
				depth--;
				if (depth === 0) return i + 1;
			}
			i++;
		}
		return null;
	}
	if (text[i] === '"') {
		i++;
		while (i < n) {
			if (text[i] === '\\') {
				i += 2;
				continue;
			}
			if (text[i] === '"') return i + 1;
			i++;
		}
		return null;
	}
	while (i < n && text[i] !== ',' && text[i] !== '}') i++;
	return i;
}

function stripBibNoise(text: string): [string, number] {
	const out: string[] = [];
	let pos = 0;
	let removed = 0;
	while (true) {
		BIB_NOISE_RE.lastIndex = pos;
		const m = BIB_NOISE_RE.exec(text);
		if (!m) {
			out.push(text.slice(pos));
			break;
		}
		const valueStart = m.index + m[0].length;
		let end = bibValueEnd(text, valueStart);
		if (end === null) {
			// Unterminated: leave the field alone rather than truncate the file.
			out.push(text.slice(pos, valueStart));
			pos = valueStart;
			continue;
		}
		out.push(text.slice(pos, m.index));
		// Take the leading separator with the field and leave the trailing
		// one, so a run of noise fields is stripped in full.
		if (m[1][0] === '{') {
			out.push('{');
			while (end < text.length && (text[end] === ' ' || text[end] === '\t')) end++;
			if (end < text.length && text[end] === ',') end++;
		}
		pos = end;
		removed++;
	}
	return [out.join(''), removed];
}

function bibOpeners(text: string): string[] {
	return [...text.matchAll(/@\s*[A-Za-z]+\s*[{(]\s*([^,\s}]+)/g)].map(m => m[1]);
}

function narrowBib(bibSrc: string, markdown: string, tmpDir: string): string {
	let text: string;
	try {
		text = fs.readFileSync(bibSrc, 'utf8');
	} catch (e) {
		return bibSrc;
	}
	const keys = citedKeys(markdown);
	if (!keys.size) return bibSrc;
	let small = filterBib(text, keys);
	if (small) {
		small = stripBibNoise(small)[0];
	} else {
		// The filter gave up; stripping alone is still worth it, verified by a
		// scan that does not share the brace matcher's possible bugs.
		const [stripped, removed] = stripBibNoise(text);
		const before = bibOpeners(text);
		const after = bibOpeners(stripped);
		if (!removed || before.length !== after.length || before.some((k, i) => k !== after[i])) return bibSrc;
		small = stripped;
	}
	if (small.length >= text.length * 0.9) return bibSrc;
	try {
		const dest = path.join(tmpDir, 'cited.bib');
		fs.writeFileSync(dest, small, 'utf8');
		return dest;
	} catch (e) {
		return bibSrc;
	}
}

// ── DOCX post-processing ───────────────────────────────────────────────

const EMPTY_HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
    </w:pPr>
  </w:p>
</w:hdr>`;

const EMPTY_FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
    </w:pPr>
  </w:p>
</w:ftr>`;

// Rough width of one character of 12pt Times and the Table style's cell
// padding, in twips. Estimates, so the algorithm errs toward giving a column
// more room than it needs.
const CHAR_TWIPS = 125;
const CELL_PAD_TWIPS = 216;
const TEXT_WIDTH_TWIPS = 9360; // US letter less 1in side margins

// pandoc sizes columns from the dash counts in the markdown separator row,
// which is a drawing of the source rather than an instruction about the
// output. Measure the content instead: each column asks for its widest cell
// (at least its longest word), and the smallest requests are filled first so
// narrow columns get their natural size and the prose column absorbs the rest.
function columnWidths(rows: string[][], total = TEXT_WIDTH_TWIPS): number[] {
	const n = rows.reduce((mx, r) => Math.max(mx, r.length), 0);
	if (!n) return [];
	const wantC = new Array(n).fill(0);
	const needC = new Array(n).fill(0);
	for (const row of rows) {
		row.slice(0, n).forEach((cell, i) => {
			const longestWord = cell.split(/\s+/).reduce((mx, w) => Math.max(mx, w.length), 1);
			needC[i] = Math.max(needC[i], longestWord);
			wantC[i] = Math.max(wantC[i], cell.length);
		});
	}
	const toTw = (c: number) => c * CHAR_TWIPS + CELL_PAD_TWIPS;
	const need = needC.map(c => toTw(Math.max(3, c)));
	const want = need.map((nd, i) => Math.max(nd, toTw(wantC[i])));

	const out = new Array(n).fill(0);
	let remaining = total;
	const left = [...Array(n).keys()].sort((a, b) => want[a] - want[b]);
	while (left.length) {
		const share = Math.floor(remaining / left.length);
		const i = left[0];
		if (want[i] > share) {
			// Everything still unassigned wants more than its share: split in
			// proportion to the asking, scaling back together if unbreakable
			// words alone would overflow the page.
			const tw = left.reduce((s, j) => s + want[j], 0) || 1;
			const alloc = new Map(left.map(j => [j, Math.max(need[j], Math.floor(remaining * want[j] / tw))]));
			const over = [...alloc.values()].reduce((s, v) => s + v, 0);
			for (const [j, v] of alloc) {
				out[j] = over > remaining ? Math.max(1, Math.floor(v * remaining / over)) : v;
			}
			remaining = 0;
			break;
		}
		out[i] = want[i];
		remaining -= want[i];
		left.shift();
	}
	// Table is 100% wide, so give any slack to the widest column.
	if (remaining > 0) out[out.indexOf(Math.max(...out))] += remaining;
	return out;
}

function cellTextRows(tableXml: string): string[][] {
	return [...tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(tr =>
		[...tr[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map(tc =>
			[...tc[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(t => t[1]).join('')));
}

function autofitTables(documentXml: string): string {
	return documentXml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, tbl => {
		const widths = columnWidths(cellTextRows(tbl));
		if (!widths.length) return tbl;
		let gi = 0;
		tbl = tbl.replace(/<w:gridCol w:w="\d+"\s*\/>/g,
			() => `<w:gridCol w:w="${widths[Math.min(gi++, widths.length - 1)]}"/>`);
		let ci = 0;
		tbl = tbl.replace(/<w:tcW w:type="dxa" w:w="\d+"\s*\/>/g,
			() => `<w:tcW w:type="dxa" w:w="${widths[ci++ % widths.length]}"/>`);
		return tbl;
	});
}

// ── UI ─────────────────────────────────────────────────────────────────

class AfterExportModal extends Modal {
	constructor(app: App, private outPath: string) {
		super(app);
	}

	onOpen() {
		const { contentEl } = this;
		const { shell } = require('electron');
		contentEl.createEl('h3', { text: 'Export complete' });
		contentEl.createEl('p', { text: this.outPath });
		new Setting(contentEl)
			.addButton(btn => btn.setButtonText('Open').setCta().onClick(async () => {
				const err = await shell.openPath(this.outPath);
				if (err) new Notice(`Could not open ${path.basename(this.outPath)}: ${err}`);
				this.close();
			}))
			.addButton(btn => btn.setButtonText('Show in folder').onClick(() => {
				shell.showItemInFolder(this.outPath);
				this.close();
			}))
			.addButton(btn => btn.setButtonText('Close').onClick(() => this.close()));
	}

	onClose() {
		this.contentEl.empty();
	}
}

class MD2PDFSettingTab extends PluginSettingTab {
	plugin: MD2PDFPlugin;

	constructor(app: App, plugin: MD2PDFPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		containerEl.createEl('h2', { text: 'MD to PDF Settings' });

		new Setting(containerEl)
			.setName('PDF Engine')
			.setDesc('Typst renders from the bundled MLA/Turabian templates and needs no LibreOffice. ' +
				'Notes whose spacing: names a reference doc other than single/double always use LibreOffice, ' +
				'as does everything when Typst is not installed.')
			.addDropdown(dropdown => dropdown
				.addOption('typst', 'Typst')
				.addOption('libreoffice', 'LibreOffice')
				.setValue(this.plugin.settings.pdfEngine)
				.onChange(async (value) => {
					this.plugin.settings.pdfEngine = value as PdfEngine;
					await this.plugin.saveSettings();
				}));

		// Default reference document dropdown
		const availableDocs = this.plugin.getAvailableRefDocs();

		new Setting(containerEl)
			.setName('Default Reference Document')
			.setDesc('Default typography/spacing template when no spacing is specified in YAML (DOCX and LibreOffice exports)')
			.addDropdown(dropdown => {
				dropdown.addOption('', '(none - use double, or first available)');
				for (const doc of availableDocs) {
					dropdown.addOption(doc, doc);
				}
				dropdown.setValue(this.plugin.settings.defaultRef);
				dropdown.onChange(async (value) => {
					this.plugin.settings.defaultRef = value;
					await this.plugin.saveSettings();
				});
			});

		// Refresh button for reference docs
		new Setting(containerEl)
			.setName('Refresh Reference Documents')
			.setDesc(availableDocs.length > 0
				? `Rescan plugin folder for .docx files — scanning ${this.plugin.getPluginFolder()}`
				: `No .docx files found in ${this.plugin.getPluginFolder()}`)
			.addButton(btn => btn
				.setButtonText('Refresh')
				.onClick(() => {
					this.display();
					new Notice('Reference documents refreshed');
				}));

		const pathSetting = (name: string, key: 'pandocPath' | 'libreOfficePath' | 'typstPath', placeholder: string) =>
			new Setting(containerEl)
				.setName(name)
				.setDesc(`Full path to the ${name.replace(' Path', '')} executable`)
				.addText(text => text
					.setPlaceholder(placeholder)
					.setValue(this.plugin.settings[key])
					.onChange(async (value) => {
						this.plugin.settings[key] = value;
						await this.plugin.saveSettings();
					}));

		pathSetting('Pandoc Path', 'pandocPath', '/usr/bin/pandoc');
		pathSetting('Typst Path', 'typstPath', '/usr/bin/typst');
		pathSetting('LibreOffice Path', 'libreOfficePath', '/usr/bin/libreoffice');

		// Auto-detect button
		new Setting(containerEl)
			.setName('Auto-detect Paths')
			.setDesc('Attempt to automatically find Pandoc, Typst and LibreOffice')
			.addButton(btn => btn
				.setButtonText('Auto-detect')
				.onClick(async () => {
					this.plugin.settings.pandocPath = '';
					this.plugin.settings.libreOfficePath = '';
					this.plugin.settings.typstPath = '';
					await this.plugin.autoDetectPaths();
					this.display();
					new Notice('Auto-detection complete');
				}));

		// Custom pandoc args
		new Setting(containerEl)
			.setName('Custom Pandoc Arguments')
			.setDesc('Additional arguments to pass to pandoc (space-separated)')
			.addText(text => text
				.setPlaceholder('--toc --number-sections')
				.setValue(this.plugin.settings.customPandocArgs)
				.onChange(async (value) => {
					this.plugin.settings.customPandocArgs = value;
					await this.plugin.saveSettings();
				}));

		// Auto-detect bibliography
		new Setting(containerEl)
			.setName('Auto-detect Bibliography')
			.setDesc('Enable citeproc when a bibliography field is found in frontmatter. Without csl:, the style: picks Chicago notes-bibliography or MLA.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.autoDetectBibliography)
				.onChange(async (value) => {
					this.plugin.settings.autoDetectBibliography = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Offer to Open Export')
			.setDesc('After exporting, offer to open the file or show it in its folder')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.offerToOpen)
				.onChange(async (value) => {
					this.plugin.settings.offerToOpen = value;
					await this.plugin.saveSettings();
				}));

		// Delete intermediate DOCX
		new Setting(containerEl)
			.setName('Delete Intermediate .docx')
			.setDesc('Delete the intermediate .docx file after a LibreOffice PDF conversion')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.deleteIntermediateDocx)
				.onChange(async (value) => {
					this.plugin.settings.deleteIntermediateDocx = value;
					await this.plugin.saveSettings();
				}));

		// PDF output location
		new Setting(containerEl)
			.setName('PDF Output Location')
			.setDesc('Where to save PDF files. Leave empty for same folder as source. Can be relative to vault (e.g., "exports/pdf") or absolute path.')
			.addText(text => text
				.setPlaceholder('(same as source file)')
				.setValue(this.plugin.settings.pdfOutputLocation)
				.onChange(async (value) => {
					this.plugin.settings.pdfOutputLocation = value;
					await this.plugin.saveSettings();
				}));

		// DOCX output location
		new Setting(containerEl)
			.setName('DOCX Output Location')
			.setDesc('Where to save .docx exports (and intermediate .docx files when kept). Leave empty for same folder as source.')
			.addText(text => text
				.setPlaceholder('(same as source file)')
				.setValue(this.plugin.settings.docxOutputLocation)
				.onChange(async (value) => {
					this.plugin.settings.docxOutputLocation = value;
					await this.plugin.saveSettings();
				}));

		// Help section - Format types
		containerEl.createEl('h3', { text: 'Style Types' });
		const formatsEl = containerEl.createEl('div');
		formatsEl.innerHTML = `
			<p>Use the <code>style:</code> YAML field to control document structure:</p>
			<ul>
				<li><code>style: chicago</code> (or <code>turabian</code>) — Turabian-style title page (title centered 1/3 down, author/course/date in bottom half)</li>
				<li><code>style: mla</code> — MLA-style header block (author, instructor, course, date left-aligned at top)</li>
				<li><em>(no style)</em> — Plain document with no special header or cover page</li>
			</ul>
			<p>Run <strong>Insert export frontmatter</strong> from the command palette to add the fields.</p>
		`;

		// Help section - Reference documents
		containerEl.createEl('h3', { text: 'Reference Documents' });
		const refEl = containerEl.createEl('div');
		refEl.innerHTML = `
			<p>Use the <code>spacing:</code> YAML field to specify typography and spacing:</p>
			<ul>
				<li><code>single</code> and <code>double</code> work with both engines</li>
				<li>Other <code>.docx</code> files in the plugin folder are used by filename without extension and always render through LibreOffice</li>
			</ul>
			<p><strong>Available reference documents:</strong> ${availableDocs.length > 0 ? availableDocs.join(', ') : '(none found)'}</p>
		`;

		// YAML examples
		containerEl.createEl('h3', { text: 'YAML Frontmatter Example' });
		const yamlEl = containerEl.createEl('div');
		yamlEl.innerHTML = `
			<pre>---
title: My Thesis Title
subtitle: An Optional Second Line
author: John Smith
course: English 626
instructor: Dr. Johnson
date: 2025-03-07
style: chicago
spacing: double
bibliography: sources/library.bib
---</pre>
		`;
	}
}
