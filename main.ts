import { App, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile, normalizePath } from 'obsidian';
import { exec, spawn } from 'child_process';
import { promisify } from 'util';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';

const execAsync = promisify(exec);

interface MD2PDFSettings {
	pandocPath: string;
	libreOfficePath: string;
	defaultRef: string;
	deleteIntermediateDocx: boolean;
	autoDetectBibliography: boolean;
	customPandocArgs: string;
	pdfOutputLocation: string;
	docxOutputLocation: string;
}

const DEFAULT_SETTINGS: MD2PDFSettings = {
	pandocPath: '',
	libreOfficePath: '',
	defaultRef: '',
	deleteIntermediateDocx: true,
	autoDetectBibliography: true,
	customPandocArgs: '',
	pdfOutputLocation: '',
	docxOutputLocation: ''
};

export default class MD2PDFPlugin extends Plugin {
	settings!: MD2PDFSettings;

	async onload() {
		await this.loadSettings();

		// Add command to convert current file
		this.addCommand({
			id: 'convert-md-to-pdf',
			name: 'Convert current Markdown to PDF',
			editorCallback: async (editor: Editor, view: MarkdownView | import('obsidian').MarkdownFileInfo) => {
				if (view.file) {
					await this.convertToPDF(view.file);
				} else {
					new Notice('No file is currently open');
				}
			}
		});

		// Add ribbon icon
		this.addRibbonIcon('file-output', 'Convert MD to PDF', async () => {
			const activeFile = this.app.workspace.getActiveFile();
			if (activeFile && activeFile.extension === 'md') {
				await this.convertToPDF(activeFile);
			} else {
				new Notice('Please open a Markdown file first');
			}
		});

		// Add settings tab
		this.addSettingTab(new MD2PDFSettingTab(this.app, this));

		// Auto-detect paths on first load
		if (!this.settings.pandocPath || !this.settings.libreOfficePath) {
			await this.autoDetectPaths();
		}
	}

	async autoDetectPaths() {
		const platform = os.platform();
		let pandocPaths: string[] = [];
		let libreofficePaths: string[] = [];

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
		}

		// Try to find pandoc
		if (!this.settings.pandocPath) {
			for (const p of pandocPaths) {
				if (fs.existsSync(p)) {
					this.settings.pandocPath = p;
					break;
				}
			}
			// Try which/where command
			if (!this.settings.pandocPath) {
				try {
					const cmd = platform === 'win32' ? 'where pandoc' : 'which pandoc';
					const { stdout } = await execAsync(cmd);
					const foundPath = stdout.trim().split('\n')[0];
					if (foundPath && fs.existsSync(foundPath)) {
						this.settings.pandocPath = foundPath;
					}
				} catch (e) {
					// Command failed, path not found
				}
			}
		}

		// Try to find LibreOffice
		if (!this.settings.libreOfficePath) {
			for (const p of libreofficePaths) {
				if (fs.existsSync(p)) {
					this.settings.libreOfficePath = p;
					break;
				}
			}
			// Try which/where command
			if (!this.settings.libreOfficePath) {
				try {
					const cmd = platform === 'win32' ? 'where soffice' : 'which libreoffice || which soffice';
					const { stdout } = await execAsync(cmd);
					const foundPath = stdout.trim().split('\n')[0];
					if (foundPath && fs.existsSync(foundPath)) {
						this.settings.libreOfficePath = foundPath;
					}
				} catch (e) {
					// Command failed, path not found
				}
			}
		}

		await this.saveSettings();
	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	// Get plugin folder path
	getPluginFolder(): string {
		const vaultPath = (this.app.vault.adapter as any).basePath;
		return path.join(vaultPath, this.app.vault.configDir, 'plugins', this.manifest.id);
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

	// Parse YAML frontmatter from content
	parseYAML(content: string): Record<string, string> {
		const yaml: Record<string, string> = {};
		const match = content.match(/^---\n([\s\S]*?)\n---/);
		if (match) {
			const lines = match[1].split('\n');
			for (const line of lines) {
				const colonIndex = line.indexOf(':');
				if (colonIndex > 0) {
					const key = line.substring(0, colonIndex).trim();
					const value = line.substring(colonIndex + 1).trim();
					yaml[key] = value;
				}
			}
		}
		return yaml;
	}

	async convertToPDF(file: TFile): Promise<void> {
		// Validate settings
		if (!this.settings.pandocPath) {
			new Notice('Pandoc path not configured. Please set it in plugin settings.');
			return;
		}
		if (!this.settings.libreOfficePath) {
			new Notice('LibreOffice path not configured. Please set it in plugin settings.');
			return;
		}
		if (!fs.existsSync(this.settings.pandocPath)) {
			new Notice(`Pandoc not found at: ${this.settings.pandocPath}`);
			return;
		}
		if (!fs.existsSync(this.settings.libreOfficePath)) {
			new Notice(`LibreOffice not found at: ${this.settings.libreOfficePath}`);
			return;
		}

		// Read file content
		const content = await this.app.vault.read(file);
		const yaml = this.parseYAML(content);
		
		// Determine style type (chicago, mla, or none)
		const formatType = yaml['style'] || '';

		// Determine reference doc
		const pluginFolder = this.getPluginFolder();
		let referenceDoc = '';

		// Check for spacing in YAML first
		if (yaml['spacing']) {
			const refPath = path.join(pluginFolder, yaml['spacing'] + '.docx');
			if (fs.existsSync(refPath)) {
				referenceDoc = refPath;
			} else {
				new Notice(`Reference doc not found: ${yaml['spacing']}.docx`);
				return;
			}
		}
		// Fall back to default ref from settings
		else if (this.settings.defaultRef) {
			const defaultRefPath = path.join(pluginFolder, this.settings.defaultRef + '.docx');
			if (fs.existsSync(defaultRefPath)) {
				referenceDoc = defaultRefPath;
			}
		}
		// Final fallback to any .docx in plugin folder
		else {
			const availableDocs = this.getAvailableRefDocs();
			if (availableDocs.length > 0) {
				referenceDoc = path.join(pluginFolder, availableDocs[0] + '.docx');
			}
		}
		
		if (!referenceDoc || !fs.existsSync(referenceDoc)) {
			new Notice('No reference.docx found. Please add a .docx file to the plugin folder.');
			return;
		}

		const formatDesc = formatType ? ` (${formatType})` : '';
		new Notice(`Converting ${file.name}${formatDesc}...`);

		try {
			// Get absolute paths
			const vaultPath = (this.app.vault.adapter as any).basePath;
			const inputPath = path.join(vaultPath, file.path);
			const dirPath = path.dirname(inputPath);
			const baseName = path.basename(file.path, '.md');
			
			// Determine output directories (use settings if provided, otherwise same as source)
			let docxOutputDir = dirPath;
			let pdfOutputDir = dirPath;
			
			if (this.settings.docxOutputLocation) {
				// Check if it's an absolute path or relative to vault
				if (path.isAbsolute(this.settings.docxOutputLocation)) {
					docxOutputDir = this.settings.docxOutputLocation;
				} else {
					docxOutputDir = path.join(vaultPath, this.settings.docxOutputLocation);
				}
				// Create directory if it doesn't exist
				if (!fs.existsSync(docxOutputDir)) {
					fs.mkdirSync(docxOutputDir, { recursive: true });
				}
			}
			
			if (this.settings.pdfOutputLocation) {
				// Check if it's an absolute path or relative to vault
				if (path.isAbsolute(this.settings.pdfOutputLocation)) {
					pdfOutputDir = this.settings.pdfOutputLocation;
				} else {
					pdfOutputDir = path.join(vaultPath, this.settings.pdfOutputLocation);
				}
				// Create directory if it doesn't exist
				if (!fs.existsSync(pdfOutputDir)) {
					fs.mkdirSync(pdfOutputDir, { recursive: true });
				}
			}
			
			const docxPath = path.join(docxOutputDir, `${baseName}.docx`);
			const pdfPath = path.join(pdfOutputDir, `${baseName}.pdf`);

			// Check for bibliography
			const hasBibliography = this.settings.autoDetectBibliography && 
				(content.includes('bibliography:') || content.includes('bibliography :'));

			// Build pandoc command
			const pandocArgs: string[] = [
				inputPath,
				'--standalone',
				'--reference-doc=' + referenceDoc
			];

			// Add citeproc if bibliography detected
			if (hasBibliography) {
				pandocArgs.push('--citeproc');
				new Notice('Bibliography detected, enabling citeproc');
			}

			// Create and add Lua filter based on format type
			const tempLuaFilter = await this.createTempLuaFilter(formatType, yaml);
			if (tempLuaFilter) {
				pandocArgs.push('--lua-filter=' + tempLuaFilter);
			}

			// Add custom args if configured
			if (this.settings.customPandocArgs) {
				const customArgs = this.settings.customPandocArgs.split(/\s+/).filter(arg => arg.length > 0);
				pandocArgs.push(...customArgs);
			}

			pandocArgs.push('-o', docxPath);

			// Run pandoc
			await this.runCommand(this.settings.pandocPath, pandocArgs, dirPath);

			// Check if DOCX was created
			if (!fs.existsSync(docxPath)) {
				throw new Error('Pandoc failed to create DOCX file');
			}

			// Strip headers or footers based on style type
			// style: mla -> keep headers, strip footers
			// style: chicago or blank -> keep footers, strip headers
			await this.stripHeadersOrFooters(docxPath, formatType);

			// Replace {{LASTNAME}} placeholder in DOCX headers/footers
			// If no author, remove the placeholder entirely; otherwise replace with last name
			if (yaml['author']) {
				const authorName = yaml['author'];
				const lastName = yaml['lastname'] || authorName.split(/\s+/).pop() || authorName;
				await this.replaceHeaderPlaceholder(docxPath, '{{LASTNAME}} ', lastName + ' ');
				await this.replaceHeaderPlaceholder(docxPath, '{{LASTNAME}}', lastName);
			} else {
				// No author - remove the placeholder (and trailing space if present)
				await this.replaceHeaderPlaceholder(docxPath, '{{LASTNAME}} ', '');
				await this.replaceHeaderPlaceholder(docxPath, '{{LASTNAME}}', '');
			}

			// Run LibreOffice to convert to PDF
			const libreOfficeArgs = [
				'--headless',
				'--convert-to', 'pdf',
				'--outdir', pdfOutputDir,
				docxPath
			];

			await this.runCommand(this.settings.libreOfficePath, libreOfficeArgs, pdfOutputDir);

			// Check if PDF was created
			if (!fs.existsSync(pdfPath)) {
				throw new Error('LibreOffice failed to create PDF file');
			}

			// Delete intermediate DOCX if configured
			if (this.settings.deleteIntermediateDocx) {
				try {
					fs.unlinkSync(docxPath);
				} catch (e) {
					console.warn('Failed to delete intermediate DOCX:', e);
				}
			}

			new Notice(`Successfully converted to ${baseName}.pdf`);

		} catch (error) {
			console.error('MD2PDF conversion error:', error);
			new Notice(`Conversion failed: ${(error as Error).message}`);
		}
	}

	async replaceHeaderPlaceholder(docxPath: string, placeholder: string, replacement: string): Promise<void> {
		const AdmZip = require('adm-zip');
		
		try {
			const zip = new AdmZip(docxPath);
			const zipEntries = zip.getEntries();
			let modified = false;
			
			for (const entry of zipEntries) {
				// Check header and footer files (word/header1.xml, word/footer1.xml, etc.)
				if (entry.entryName.match(/word\/(header|footer)\d*\.xml/)) {
					let content = entry.getData().toString('utf8');
					if (content.includes(placeholder)) {
						content = content.split(placeholder).join(replacement);
						zip.updateFile(entry.entryName, Buffer.from(content, 'utf8'));
						modified = true;
					}
				}
			}
			
			if (modified) {
				zip.writeZip(docxPath);
			}
		} catch (e) {
			console.warn('Failed to replace header placeholder:', e);
			// Non-fatal - continue with conversion
		}
	}

	// Strip headers or footers from DOCX based on style type
	// style: mla -> keep headers, strip footers
	// style: chicago or blank -> keep footers, strip headers
	async stripHeadersOrFooters(docxPath: string, formatType: string): Promise<void> {
		const AdmZip = require('adm-zip');

		// Determine what to strip
		const stripHeaders = formatType !== 'mla'; // strip headers for chicago or blank
		const stripFooters = formatType === 'mla'; // strip footers only for mla format
		
		try {
			const zip = new AdmZip(docxPath);
			const zipEntries = zip.getEntries();
			let modified = false;
			
			for (const entry of zipEntries) {
				// Check for header files
				if (stripHeaders && entry.entryName.match(/word\/header\d*\.xml/)) {
					// Replace content with empty header
					const emptyHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
    </w:pPr>
  </w:p>
</w:hdr>`;
					zip.updateFile(entry.entryName, Buffer.from(emptyHeader, 'utf8'));
					modified = true;
				}
				
				// Check for footer files
				if (stripFooters && entry.entryName.match(/word\/footer\d*\.xml/)) {
					// Replace content with empty footer
					const emptyFooter = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
    </w:pPr>
  </w:p>
</w:ftr>`;
					zip.updateFile(entry.entryName, Buffer.from(emptyFooter, 'utf8'));
					modified = true;
				}
			}
			
			if (modified) {
				zip.writeZip(docxPath);
			}
		} catch (e) {
			console.warn('Failed to strip headers/footers:', e);
			// Non-fatal - continue with conversion
		}
	}

	async createTempLuaFilter(formatType: string, yaml: Record<string, string>): Promise<string | null> {
		try {
			const tempDir = os.tmpdir();
			const filterPath = path.join(tempDir, 'md2pdf-pagebreak.lua');
			
			let luaContent: string;
			
			if (formatType === 'mla') {
				// MLA-style header format
				luaContent = this.generateHeaderFilter(yaml);
			} else if (formatType === 'chicago') {
				// Cover page format (Turabian/Chicago style)
				luaContent = this.generateCoverPageFilter(yaml);
			} else {
				// No format specified - basic filter (just bibliography page break)
				luaContent = this.generateBasicFilter();
			}

			fs.writeFileSync(filterPath, luaContent, 'utf8');
			return filterPath;
		} catch (e) {
			console.warn('Failed to create temp Lua filter:', e);
			return null;
		}
	}

	luaBibHelpers(): string {
		return `local function escape_xml(s)
  s = s:gsub("&", "&amp;")
  s = s:gsub("<", "&lt;")
  s = s:gsub(">", "&gt;")
  return s
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
  local text = pandoc.utils.stringify(block)
  return text:match("Bibliography") or text:match("References") or text:match("Works Cited")
end`;
	}

	generateBasicFilter(): string {
		return `${this.luaBibHelpers()}
function Pandoc(doc)
  local new_blocks = {}
  local in_bib = false
  for i, block in ipairs(doc.blocks) do
    if is_bib_heading(block) then
      in_bib = true
      table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading%d"/>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], block.level, pandoc.utils.stringify(block))))
    elseif in_bib and block.t == "Header" then
      in_bib = false
      table.insert(new_blocks, block)
    elseif in_bib and block.t == "Para" then
      table.insert(new_blocks, bib_entry_block(block))
    else
      table.insert(new_blocks, block)
    end
  end
  doc.blocks = new_blocks
  return doc
end`;
	}

	generateCoverPageFilter(yaml: Record<string, string>): string {
		// Extract YAML values for the Lua filter
		const title = (yaml['title'] || '').replace(/"/g, '\\"');
		const author = (yaml['author'] || '').replace(/"/g, '\\"');
		const course = (yaml['course'] || '').replace(/"/g, '\\"');
		const instructor = (yaml['instructor'] || '').replace(/"/g, '\\"');
		const date = (yaml['date'] || '').replace(/"/g, '\\"');
		
		return `${this.luaBibHelpers()}
-- Cover page format (Turabian style)
-- Title centered 1/3 down the page
-- Author, course, date centered in bottom half

local meta_title = "${title}"
local meta_author = "${author}"
local meta_course = "${course}"
local meta_instructor = "${instructor}"
local meta_date = "${date}"

-- Format date from YYYY-MM-DD to "Month DD, YYYY"
local function format_date(date_str)
  local months = {
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  }
  local year, month, day = date_str:match("(%d+)-(%d+)-(%d+)")
  if year and month and day then
    local month_name = months[tonumber(month)]
    if month_name then
      return string.format("%s %d, %s", month_name, tonumber(day), year)
    end
  end
  return date_str
end

function Header(el)
  local text = pandoc.utils.stringify(el)
  if text:match("Bibliography") or text:match("References") or text:match("Works Cited") then
    return pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading%d"/>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], el.level, text))
  end
  return el
end

function Meta(meta)
  -- Capture from document if not in YAML passed to filter
  if meta.title and meta_title == "" then
    meta_title = pandoc.utils.stringify(meta.title)
  end
  if meta.author and meta_author == "" then
    meta_author = pandoc.utils.stringify(meta.author)
  end
  if meta.course and meta_course == "" then
    meta_course = pandoc.utils.stringify(meta.course)
  end
  if meta.instructor and meta_instructor == "" then
    meta_instructor = pandoc.utils.stringify(meta.instructor)
  end
  if meta.date and meta_date == "" then
    meta_date = pandoc.utils.stringify(meta.date)
  end
  -- Remove from default rendering
  meta.author = nil
  meta.date = nil
  meta.title = nil
  meta.course = nil
  meta.instructor = nil
  return meta
end

function Pandoc(doc)
  local new_blocks = {}
  
  -- Count how many info lines we'll have (author, course, instructor, date)
  local info_count = 0
  if meta_author and meta_author ~= "" then info_count = info_count + 1 end
  if meta_course and meta_course ~= "" then info_count = info_count + 1 end
  if meta_instructor and meta_instructor ~= "" then info_count = info_count + 1 end
  if meta_date and meta_date ~= "" then info_count = info_count + 1 end
  
  -- Title positioned about 1/3 down (using pt spacing instead of empty paragraphs)
  -- 2400 twips = ~1.67 inches from top (reduced from 3168)
  if meta_title and meta_title ~= "" then
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:before="2400" w:after="0" w:line="480" w:lineRule="auto"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], meta_title)))
  end
  
  -- Space between title and author info
  -- Position author block roughly in bottom third of page
  -- 4320 twips = ~3 inches gap (reduced from 5760)
  local gap_before_author = 4320
  local first_info = true
  
  -- Add centered author
  if meta_author and meta_author ~= "" then
    local spacing_before = first_info and gap_before_author or 0
    first_info = false
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:before="%d" w:after="0" w:line="480" w:lineRule="auto"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], spacing_before, meta_author)))
  end
  
  -- Add centered course
  if meta_course and meta_course ~= "" then
    local spacing_before = first_info and gap_before_author or 0
    first_info = false
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:before="%d" w:after="0" w:line="480" w:lineRule="auto"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], spacing_before, meta_course)))
  end
  
  -- Add centered instructor (if present)
  if meta_instructor and meta_instructor ~= "" then
    local spacing_before = first_info and gap_before_author or 0
    first_info = false
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:before="%d" w:after="0" w:line="480" w:lineRule="auto"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], spacing_before, meta_instructor)))
  end
  
  -- Add centered date
  if meta_date and meta_date ~= "" then
    local formatted_date = format_date(meta_date)
    local spacing_before = first_info and gap_before_author or 0
    first_info = false
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:before="%d" w:after="0" w:line="480" w:lineRule="auto"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], spacing_before, formatted_date)))
  end
  
  -- Add all original content blocks, with page break before the first one
  local page_break_inserted = false
  local in_bib = false
  for i, block in ipairs(doc.blocks) do
    if is_bib_heading(block) then
      in_bib = true
      if not page_break_inserted then
        page_break_inserted = true
      end
      table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading%d"/>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], block.level, pandoc.utils.stringify(block))))
    elseif in_bib and block.t == "Header" then
      in_bib = false
      table.insert(new_blocks, block)
    elseif in_bib and block.t == "Para" then
      table.insert(new_blocks, bib_entry_block(block))
    else
      if not page_break_inserted then
        -- Insert page break before first real content block
        if block.t == "Header" or
           (block.t == "Para" and #block.content > 0) or
           block.t == "CodeBlock" or
           block.t == "BulletList" or
           block.t == "OrderedList" or
           block.t == "Table" or
           block.t == "BlockQuote" or
           block.t == "RawBlock" then
          table.insert(new_blocks, pandoc.RawBlock('openxml', [[
<w:p>
  <w:pPr>
    <w:pageBreakBefore/>
  </w:pPr>
</w:p>]]))
          page_break_inserted = true
        end
      end
      table.insert(new_blocks, block)
    end
  end

  -- If no content blocks were found, still need to close out properly
  if not page_break_inserted then
    table.insert(new_blocks, pandoc.RawBlock('openxml', [[
<w:p>
  <w:pPr>
    <w:pageBreakBefore/>
  </w:pPr>
</w:p>]]))
  end
  
  doc.blocks = new_blocks
  return doc
end`;
	}

	generateHeaderFilter(yaml: Record<string, string>): string {
		// Extract YAML values for the Lua filter
		const title = (yaml['title'] || '').replace(/"/g, '\\"');
		const author = (yaml['author'] || '').replace(/"/g, '\\"');
		const course = (yaml['course'] || '').replace(/"/g, '\\"');
		const instructor = (yaml['instructor'] || '').replace(/"/g, '\\"');
		const date = (yaml['date'] || '').replace(/"/g, '\\"');
		
		return `${this.luaBibHelpers()}
-- MLA Header format
-- Page 1: Info block (Author, Instructor, Course, Date) left-aligned + centered title
-- Every page: Running header "LastName PageNumber" in top right (handled by reference.docx)

local meta_title = "${title}"
local meta_author = "${author}"
local meta_course = "${course}"
local meta_instructor = "${instructor}"
local meta_date = "${date}"

-- Format date from YYYY-MM-DD to "DD Month YYYY" (MLA style)
local function format_date(date_str)
  local months = {
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  }
  local year, month, day = date_str:match("(%d+)-(%d+)-(%d+)")
  if year and month and day then
    local month_name = months[tonumber(month)]
    if month_name then
      return string.format("%d %s %s", tonumber(day), month_name, year)
    end
  end
  return date_str
end

function Header(el)
  local text = pandoc.utils.stringify(el)
  if text:match("Bibliography") or text:match("References") or text:match("Works Cited") then
    return pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading%d"/>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], el.level, text))
  end
  return el
end

function Meta(meta)
  -- Capture from document if not in YAML passed to filter
  if meta.title and meta_title == "" then
    meta_title = pandoc.utils.stringify(meta.title)
  end
  if meta.author and meta_author == "" then
    meta_author = pandoc.utils.stringify(meta.author)
  end
  if meta.course and meta_course == "" then
    meta_course = pandoc.utils.stringify(meta.course)
  end
  if meta.instructor and meta_instructor == "" then
    meta_instructor = pandoc.utils.stringify(meta.instructor)
  end
  if meta.date and meta_date == "" then
    meta_date = pandoc.utils.stringify(meta.date)
  end
  -- Remove from default rendering
  meta.author = nil
  meta.date = nil
  meta.title = nil
  meta.course = nil
  meta.instructor = nil
  return meta
end

function Pandoc(doc)
  local new_blocks = {}
  
  -- Build the MLA header block (Author, Instructor, Course, Date - each on own line)
  
  -- Author line
  if meta_author and meta_author ~= "" then
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:after="0" w:line="480" w:lineRule="auto"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], meta_author)))
  end
  
  -- Instructor line
  if meta_instructor and meta_instructor ~= "" then
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:after="0" w:line="480" w:lineRule="auto"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], meta_instructor)))
  end
  
  -- Course line
  if meta_course and meta_course ~= "" then
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:after="0" w:line="480" w:lineRule="auto"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], meta_course)))
  end
  
  -- Date line
  if meta_date and meta_date ~= "" then
    local formatted_date = format_date(meta_date)
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:after="0" w:line="480" w:lineRule="auto"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], formatted_date)))
  end
  
  -- Add centered title
  if meta_title and meta_title ~= "" then
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:after="0" w:line="480" w:lineRule="auto"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], meta_title)))
  end
  
  -- Add all original content blocks
  local in_bib = false
  for i, block in ipairs(doc.blocks) do
    if is_bib_heading(block) then
      in_bib = true
      table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading%d"/>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t>%s</w:t>
  </w:r>
</w:p>]], block.level, pandoc.utils.stringify(block))))
    elseif in_bib and block.t == "Header" then
      in_bib = false
      table.insert(new_blocks, block)
    elseif in_bib and block.t == "Para" then
      table.insert(new_blocks, bib_entry_block(block))
    else
      table.insert(new_blocks, block)
    end
  end
  
  doc.blocks = new_blocks
  return doc
end`;
	}

	runCommand(command: string, args: string[], cwd: string): Promise<void> {
		return new Promise((resolve, reject) => {
			const isWindows = os.platform() === 'win32';
			
			let proc;
			if (isWindows) {
				// On Windows, quote the command and use shell
				const quotedCommand = `"${command}"`;
				const quotedArgs = args.map(arg => {
					// Quote args that contain spaces
					if (arg.includes(' ')) {
						return `"${arg}"`;
					}
					return arg;
				});
				proc = spawn(quotedCommand, quotedArgs, {
					cwd: cwd,
					shell: true
				});
			} else {
				proc = spawn(command, args, {
					cwd: cwd
				});
			}

			let stderr = '';
			let stdout = '';

			proc.stdout.on('data', (data) => {
				stdout += data.toString();
			});

			proc.stderr.on('data', (data) => {
				stderr += data.toString();
			});

			proc.on('close', (code) => {
				if (code === 0) {
					resolve();
				} else {
					reject(new Error(`Command failed with code ${code}: ${stderr || stdout}`));
				}
			});

			proc.on('error', (err) => {
				reject(err);
			});
		});
	}

	onunload() {
		// Cleanup if needed
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

		// Default reference document dropdown
		const availableDocs = this.plugin.getAvailableRefDocs();
		
		new Setting(containerEl)
			.setName('Default Reference Document')
			.setDesc('Default typography/spacing template when no spacing is specified in YAML')
			.addDropdown(dropdown => {
				dropdown.addOption('', '(none - use first available)');
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
			.setDesc('Rescan plugin folder for .docx files')
			.addButton(btn => btn
				.setButtonText('Refresh')
				.onClick(() => {
					this.display();
					new Notice('Reference documents refreshed');
				}));

		// Pandoc path
		new Setting(containerEl)
			.setName('Pandoc Path')
			.setDesc('Full path to pandoc executable')
			.addText(text => text
				.setPlaceholder('/usr/bin/pandoc')
				.setValue(this.plugin.settings.pandocPath)
				.onChange(async (value) => {
					this.plugin.settings.pandocPath = value;
					await this.plugin.saveSettings();
				}));

		// LibreOffice path
		new Setting(containerEl)
			.setName('LibreOffice Path')
			.setDesc('Full path to LibreOffice/soffice executable')
			.addText(text => text
				.setPlaceholder('/usr/bin/libreoffice')
				.setValue(this.plugin.settings.libreOfficePath)
				.onChange(async (value) => {
					this.plugin.settings.libreOfficePath = value;
					await this.plugin.saveSettings();
				}));

		// Auto-detect button
		new Setting(containerEl)
			.setName('Auto-detect Paths')
			.setDesc('Attempt to automatically find Pandoc and LibreOffice')
			.addButton(btn => btn
				.setButtonText('Auto-detect')
				.onClick(async () => {
					this.plugin.settings.pandocPath = '';
					this.plugin.settings.libreOfficePath = '';
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
			.setDesc('Automatically enable citeproc when bibliography field is found in frontmatter')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.autoDetectBibliography)
				.onChange(async (value) => {
					this.plugin.settings.autoDetectBibliography = value;
					await this.plugin.saveSettings();
				}));

		// Delete intermediate DOCX
		new Setting(containerEl)
			.setName('Delete Intermediate .docx')
			.setDesc('Delete the intermediate .docx file after PDF conversion')
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
			.setDesc('Where to save intermediate .docx files (when not deleted). Leave empty for same folder as source.')
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
				<li><code>style: chicago</code> — Turabian-style title page (title centered 1/3 down, author/course/date in bottom half)</li>
				<li><code>style: mla</code> — MLA-style header block (author, instructor, course, date left-aligned at top)</li>
				<li><em>(no style)</em> — Plain document with no special header or cover page</li>
			</ul>
		`;

		// Help section - Reference documents
		containerEl.createEl('h3', { text: 'Reference Documents' });
		const refEl = containerEl.createEl('div');
		refEl.innerHTML = `
			<p>Use the <code>spacing:</code> YAML field to specify typography and spacing:</p>
			<ul>
				<li>Place <code>.docx</code> files in the plugin folder (<code>.obsidian/plugins/md2pdf/</code>)</li>
				<li>Reference by filename without extension: <code>spacing: double</code> → uses <code>double.docx</code></li>
				<li>The reference document controls fonts, margins, line spacing, etc.</li>
			</ul>
			<p><strong>Available reference documents:</strong> ${availableDocs.length > 0 ? availableDocs.join(', ') : '(none found)'}</p>
		`;

		// YAML examples
		containerEl.createEl('h3', { text: 'YAML Frontmatter Examples' });
		const yamlEl = containerEl.createEl('div');
		yamlEl.innerHTML = `
			<p>Cover page (Turabian/Chicago style):</p>
			<pre>---
title: My Thesis Title
author: John Smith
course: English 626
instructor: Dr. Johnson
date: 2025-03-07
style: chicago
spacing: double
---</pre>
			<p>MLA header style:</p>
			<pre>---
title: Literary Analysis
author: Jane Doe
instructor: Mrs. Williams
course: 12th Grade Literature
date: 2025-03-07
style: mla
spacing: double
---</pre>
			<p>Plain document (no header/cover):</p>
			<pre>---
title: Quick Notes
spacing: single
---</pre>
		`;
	}
}
