// docx.js — minimal, dependency-free OOXML (.docx) writer.
//
// Produces a WordprocessingML package and zips it into a .docx using .NET's
// System.IO.Compression (invoked via PowerShell). No Word, Python, or npm
// packages required — only Node core + Windows PowerShell/.NET.
//
// The document is a single "continuous" section laid out in two columns, so
// Word flows body content column-to-column and onto as many pages as needed.

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// A "run" is a span of text with optional bold. Text is preserved verbatim
// (leading/trailing spaces kept) via xml:space="preserve".
function run(text, { bold = false } = {}) {
  const rpr = bold ? '<w:rPr><w:b/></w:rPr>' : '';
  return `<w:r>${rpr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

// Build a paragraph from an array of runs (or a single string).
// opts: { style, spaceBefore, spaceAfter, keepNext, outlineLevel }
function para(runs, opts = {}) {
  const list = Array.isArray(runs) ? runs : [run(runs)];
  const pprParts = [];
  if (opts.style) pprParts.push(`<w:pStyle w:val="${opts.style}"/>`);
  if (opts.keepNext) pprParts.push('<w:keepNext/>');
  if (opts.outlineLevel != null) pprParts.push(`<w:outlineLvl w:val="${opts.outlineLevel}"/>`);
  const spacingAttrs = [];
  if (opts.spaceBefore != null) spacingAttrs.push(`w:before="${opts.spaceBefore}"`);
  if (opts.spaceAfter != null) spacingAttrs.push(`w:after="${opts.spaceAfter}"`);
  if (spacingAttrs.length) pprParts.push(`<w:spacing ${spacingAttrs.join(' ')}/>`);
  const ppr = pprParts.length ? `<w:pPr>${pprParts.join('')}</w:pPr>` : '';
  return `<w:p>${ppr}${list.join('')}</w:p>`;
}

function heading(text, level = 1) {
  return para([run(text, { bold: true })], {
    style: `Heading${level}`,
    outlineLevel: level - 1,
    spaceBefore: 160,
    spaceAfter: 60,
    keepNext: true,
  });
}

// A hard column break (forces the following content into the next column).
function columnBreak() {
  return '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
}

// A hard page break.
function pageBreak() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

// Section properties: N columns, US Letter portrait. `type` is the section
// start type ('continuous' or 'nextPage').
function sectPrXml(numCols, type) {
  return `<w:sectPr>
<w:type w:val="${type}"/>
<w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1080" w:bottom="1080" w:left="1080" w:right="1080" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:num="${numCols}" w:space="720"/>
<w:docGrid w:linePitch="360"/>
</w:sectPr>`;
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="40" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="60"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="40"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="80" w:after="20"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
</w:styles>`;

function coreProps(title) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title>${esc(title)}</dc:title>
<dc:creator>Pack 127 Calendar Generator</dc:creator>
</cp:coreProperties>`;
}

const APP_PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>PackCalendar generator</Application>
</Properties>`;

// Assemble body pieces into a document.xml from an array of sections.
// Each section: { blocks: string[], numCols: number, type: 'continuous'|'nextPage' }.
// Every section except the last ends with a paragraph carrying its <w:sectPr>;
// the final section's <w:sectPr> is a direct child of <w:body>.
function buildDocumentXml(sections) {
  const parts = [];
  sections.forEach((sec, i) => {
    const isLast = i === sections.length - 1;
    parts.push(sec.blocks.join('\n'));
    if (!isLast) {
      parts.push(`<w:p><w:pPr>${sectPrXml(sec.numCols, sec.type)}</w:pPr></w:p>`);
    }
  });
  const lastSec = sections[sections.length - 1];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
${parts.join('\n')}
${sectPrXml(lastSec.numCols, lastSec.type)}
</w:body>
</w:document>`;
}

// Write the package to disk and zip into `outPath` (.docx).
// `sections` is an array of { blocks, numCols, type } (see buildDocumentXml).
function writeDocx(outPath, { title, sections }) {
  const build = fs.mkdtempSync(path.join(os.tmpdir(), 'packcal-'));
  try {
    fs.mkdirSync(path.join(build, '_rels'));
    fs.mkdirSync(path.join(build, 'word', '_rels'), { recursive: true });
    fs.mkdirSync(path.join(build, 'docProps'));
    fs.writeFileSync(path.join(build, '[Content_Types].xml'), CONTENT_TYPES);
    fs.writeFileSync(path.join(build, '_rels', '.rels'), ROOT_RELS);
    fs.writeFileSync(path.join(build, 'docProps', 'core.xml'), coreProps(title));
    fs.writeFileSync(path.join(build, 'docProps', 'app.xml'), APP_PROPS);
    fs.writeFileSync(path.join(build, 'word', 'styles.xml'), STYLES);
    fs.writeFileSync(path.join(build, 'word', '_rels', 'document.xml.rels'), DOC_RELS);
    fs.writeFileSync(path.join(build, 'word', 'document.xml'), buildDocumentXml(sections));

    const abs = path.resolve(outPath);
    fs.rmSync(abs, { force: true });
    if (process.platform === 'win32') {
      const q = (s) => s.replace(/'/g, "''");
      const psCmd =
        'Add-Type -AssemblyName System.IO.Compression.FileSystem; ' +
        `[System.IO.Compression.ZipFile]::CreateFromDirectory('${q(build)}','${q(abs)}')`;
      execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', psCmd]);
    } else {
      // Linux/macOS (e.g. GitHub Actions runners): use the `zip` CLI. Zip from
      // within the build dir so entry paths are relative to the package root,
      // as the .docx (OOXML) package requires.
      execFileSync('zip', ['-r', '-X', '-q', abs, '.'], { cwd: build });
    }
  } finally {
    fs.rmSync(build, { recursive: true, force: true });
  }
  return outPath;
}

module.exports = {
  esc, run, para, heading, columnBreak, pageBreak, writeDocx,
};
