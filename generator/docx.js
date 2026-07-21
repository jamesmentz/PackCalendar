// docx.js — minimal, dependency-free OOXML (.docx) writer.
//
// Produces a WordprocessingML package and zips it into a .docx using a small
// built-in, fully deterministic ZIP writer (pure Node core — no PowerShell,
// no `zip` CLI, no npm). Determinism matters: the same input always yields
// byte-identical output on every platform and Node version, so CI regenerating
// the calendar produces no spurious diff/commit when nothing substantive changed.
//
// The document is a single "continuous" section laid out in two columns, so
// Word flows body content column-to-column and onto as many pages as needed.

const fs = require('fs');
const path = require('path');

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

// --- Deterministic ZIP writer -----------------------------------------------
// A .docx is just a ZIP of the package parts. We build it by hand so the bytes
// are reproducible: every entry is STORED (no compression, so output never
// depends on the bundled zlib version), timestamps are pinned to the ZIP epoch
// (1980-01-01), entry names use forward slashes, and the entry order is fixed.

const CRC32_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = ~0;
  for (let i = 0; i < buf.length; i++) c = CRC32_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (~c) >>> 0;
}

const DOS_TIME = 0;       // 00:00:00
const DOS_DATE = 0x0021;  // 1980-01-01 (ZIP epoch)

// entries: [{ name, data: Buffer }] — returns the complete ZIP as a Buffer.
function buildZip(entries) {
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;
  for (const { name, data } of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(data);
    const size = data.length;

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); // local file header signature
    local.writeUInt16LE(20, 4);         // version needed to extract
    local.writeUInt16LE(0, 6);          // general purpose bit flag
    local.writeUInt16LE(0, 8);          // compression method: 0 = stored
    local.writeUInt16LE(DOS_TIME, 10);
    local.writeUInt16LE(DOS_DATE, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(size, 18);      // compressed size
    local.writeUInt32LE(size, 22);      // uncompressed size
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);         // extra field length
    localChunks.push(local, nameBuf, data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); // central directory header signature
    central.writeUInt16LE(20, 4);         // version made by
    central.writeUInt16LE(20, 6);         // version needed to extract
    central.writeUInt16LE(0, 8);          // general purpose bit flag
    central.writeUInt16LE(0, 10);         // compression method
    central.writeUInt16LE(DOS_TIME, 12);
    central.writeUInt16LE(DOS_DATE, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(size, 20);      // compressed size
    central.writeUInt32LE(size, 24);      // uncompressed size
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);         // extra field length
    central.writeUInt16LE(0, 32);         // file comment length
    central.writeUInt16LE(0, 34);         // disk number start
    central.writeUInt16LE(0, 36);         // internal file attributes
    central.writeUInt32LE(0, 38);         // external file attributes
    central.writeUInt32LE(offset, 42);    // relative offset of local header
    centralChunks.push(central, nameBuf);

    offset += local.length + nameBuf.length + data.length;
  }

  const localBuf = Buffer.concat(localChunks);
  const centralBuf = Buffer.concat(centralChunks);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);          // end of central directory signature
  eocd.writeUInt16LE(0, 4);                   // number of this disk
  eocd.writeUInt16LE(0, 6);                   // disk with start of central dir
  eocd.writeUInt16LE(entries.length, 8);      // central dir records on this disk
  eocd.writeUInt16LE(entries.length, 10);     // total central dir records
  eocd.writeUInt32LE(centralBuf.length, 12);  // size of central directory
  eocd.writeUInt32LE(localBuf.length, 16);    // offset of central directory
  eocd.writeUInt16LE(0, 20);                  // .zip file comment length
  return Buffer.concat([localBuf, centralBuf, eocd]);
}

// Write the package into `outPath` (.docx) as a deterministic ZIP.
// `sections` is an array of { blocks, numCols, type } (see buildDocumentXml).
function writeDocx(outPath, { title, sections }) {
  const b = (s) => Buffer.from(s, 'utf8');
  // Fixed entry order; [Content_Types].xml first per the OOXML convention.
  const entries = [
    { name: '[Content_Types].xml', data: b(CONTENT_TYPES) },
    { name: '_rels/.rels', data: b(ROOT_RELS) },
    { name: 'docProps/core.xml', data: b(coreProps(title)) },
    { name: 'docProps/app.xml', data: b(APP_PROPS) },
    { name: 'word/styles.xml', data: b(STYLES) },
    { name: 'word/_rels/document.xml.rels', data: b(DOC_RELS) },
    { name: 'word/document.xml', data: b(buildDocumentXml(sections)) },
  ];
  const abs = path.resolve(outPath);
  fs.writeFileSync(abs, buildZip(entries));
  return outPath;
}

module.exports = {
  esc, run, para, heading, columnBreak, pageBreak, writeDocx,
};
