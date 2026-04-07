const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

function escapeXml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function escapePowerShell(value) {
    return String(value).replace(/'/g, "''");
}

function readArgs() {
    const args = process.argv.slice(2);
    const inputArg = args[0];
    const outputArg = args[1];
    let packageDir = null;

    for (let i = 2; i < args.length; i++) {
        if (args[i] === "--package-dir") {
            packageDir = args[i + 1] ? path.resolve(args[i + 1]) : null;
            i++;
        }
    }

    if (!inputArg || !outputArg) {
        throw new Error(
            "Usage: node scripts/export-markdown-to-docx.js <input.md> <output.docx> [--package-dir <dir>]"
        );
    }

    return {
        inputPath: path.resolve(inputArg),
        outputPath: path.resolve(outputArg),
        packageDir,
    };
}

function splitParagraphRuns(text) {
    const runs = [];
    const tokenRegex = /(\*\*[^*]+\*\*|`[^`]+`)/g;
    let lastIndex = 0;
    let match;

    while ((match = tokenRegex.exec(text)) !== null) {
        if (match.index > lastIndex) {
            runs.push({ text: text.slice(lastIndex, match.index), kind: "text" });
        }

        const token = match[0];
        if (token.startsWith("**")) {
            runs.push({ text: token.slice(2, -2), kind: "bold" });
        } else if (token.startsWith("`")) {
            runs.push({ text: token.slice(1, -1), kind: "code" });
        }

        lastIndex = match.index + token.length;
    }

    if (lastIndex < text.length) {
        runs.push({ text: text.slice(lastIndex), kind: "text" });
    }

    return runs;
}

function buildRunsXml(text) {
    const runs = splitParagraphRuns(text);
    if (!runs.length) {
        return `<w:r><w:t xml:space="preserve"></w:t></w:r>`;
    }

    return runs.map(run => {
        let runProps = "";
        if (run.kind === "bold") {
            runProps = "<w:rPr><w:b/></w:rPr>";
        } else if (run.kind === "code") {
            runProps = "<w:rPr><w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\"/><w:sz w:val=\"20\"/></w:rPr>";
        }

        return `<w:r>${runProps}<w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`;
    }).join("");
}

function paragraphXml(text, styleId = null) {
    const styleXml = styleId ? `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` : "";
    return `<w:p>${styleXml}${buildRunsXml(text)}</w:p>`;
}

function codeParagraphXml(text) {
    return `<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function parseMarkdown(markdown) {
    const lines = markdown.replace(/\r\n/g, "\n").split("\n");
    const blocks = [];
    let paragraphBuffer = [];
    let inCodeBlock = false;

    function flushParagraph() {
        const text = paragraphBuffer.join(" ").trim();
        if (text) {
            blocks.push({ type: "paragraph", text });
        }
        paragraphBuffer = [];
    }

    for (const line of lines) {
        if (line.startsWith("```")) {
            flushParagraph();
            inCodeBlock = !inCodeBlock;
            continue;
        }

        if (inCodeBlock) {
            blocks.push({ type: "code", text: line });
            continue;
        }

        if (/^###\s+/.test(line)) {
            flushParagraph();
            blocks.push({ type: "heading3", text: line.replace(/^###\s+/, "").trim() });
            continue;
        }

        if (/^##\s+/.test(line)) {
            flushParagraph();
            blocks.push({ type: "heading2", text: line.replace(/^##\s+/, "").trim() });
            continue;
        }

        if (/^#\s+/.test(line)) {
            flushParagraph();
            blocks.push({ type: "heading1", text: line.replace(/^#\s+/, "").trim() });
            continue;
        }

        if (/^-\s+/.test(line)) {
            flushParagraph();
            blocks.push({ type: "bullet", text: line.replace(/^-\s+/, "").trim() });
            continue;
        }

        if (!line.trim()) {
            flushParagraph();
            continue;
        }

        paragraphBuffer.push(line.trim());
    }

    flushParagraph();
    return blocks;
}

function buildDocumentXml(blocks) {
    const body = [];
    let usedTitle = false;

    for (const block of blocks) {
        if (block.type === "heading1") {
            body.push(paragraphXml(block.text, usedTitle ? "Heading1" : "Title"));
            usedTitle = true;
            continue;
        }

        if (block.type === "heading2") {
            body.push(paragraphXml(block.text, "Heading1"));
            continue;
        }

        if (block.type === "heading3") {
            body.push(paragraphXml(block.text, "Heading2"));
            continue;
        }

        if (block.type === "bullet") {
            body.push(paragraphXml(`${String.fromCharCode(8226)} ${block.text}`, "ListParagraph"));
            continue;
        }

        if (block.type === "code") {
            body.push(codeParagraphXml(block.text));
            continue;
        }

        body.push(paragraphXml(block.text));
    }

    body.push(
        "<w:sectPr>" +
        "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
        "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
        "</w:sectPr>"
    );

    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" " +
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
        "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" " +
        "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
        "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " +
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
        "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" " +
        "xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" " +
        "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 wp14\">" +
        `<w:body>${body.join("")}</w:body>` +
        "</w:document>"
    );
}

function buildStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="240"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="240" w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="180" w:after="80"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="360"/>
      <w:spacing w:after="60"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CodeBlock">
    <w:name w:val="Code Block"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="40" w:after="40"/>
      <w:ind w:left="360" w:right="360"/>
      <w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>
      <w:sz w:val="18"/>
    </w:rPr>
  </w:style>
</w:styles>`;
}

function buildContentTypesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function buildRootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function buildDocumentRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function buildCoreXml(title) {
    const created = new Date().toISOString();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(title)}</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>
</cp:coreProperties>`;
}

function buildAppXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>OpenAI Codex</Application>
</Properties>`;
}

function ensureDir(dirPath) {
    fs.mkdirSync(dirPath, { recursive: true });
}

function writePackage(tempDir, title, documentXml, stylesXml) {
    ensureDir(path.join(tempDir, "_rels"));
    ensureDir(path.join(tempDir, "word", "_rels"));
    ensureDir(path.join(tempDir, "docProps"));

    fs.writeFileSync(path.join(tempDir, "[Content_Types].xml"), buildContentTypesXml(), "utf8");
    fs.writeFileSync(path.join(tempDir, "_rels", ".rels"), buildRootRelsXml(), "utf8");
    fs.writeFileSync(path.join(tempDir, "word", "document.xml"), documentXml, "utf8");
    fs.writeFileSync(path.join(tempDir, "word", "styles.xml"), stylesXml, "utf8");
    fs.writeFileSync(path.join(tempDir, "word", "_rels", "document.xml.rels"), buildDocumentRelsXml(), "utf8");
    fs.writeFileSync(path.join(tempDir, "docProps", "core.xml"), buildCoreXml(title), "utf8");
    fs.writeFileSync(path.join(tempDir, "docProps", "app.xml"), buildAppXml(), "utf8");
}

function createZipFromDir(sourceDir, zipPath) {
    const command =
        "Add-Type -AssemblyName System.IO.Compression.FileSystem; " +
        `$src='${escapePowerShell(sourceDir)}'; ` +
        `$zip='${escapePowerShell(zipPath)}'; ` +
        "if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force; } " +
        "[System.IO.Compression.ZipFile]::CreateFromDirectory($src, $zip)";
    const powershellPath = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
    const result = spawnSync(powershellPath, ["-NoProfile", "-Command", command], {
        encoding: "utf8",
    });

    if (result.status !== 0) {
        throw new Error(
            `Zip creation failed. status=${result.status} signal=${result.signal || ""} ` +
            `error=${result.error ? result.error.message : ""} stdout=${result.stdout || ""} stderr=${result.stderr || ""}`
        );
    }
}

function main() {
    const { inputPath, outputPath, packageDir } = readArgs();
    const markdown = fs.readFileSync(inputPath, "utf8");
    const blocks = parseMarkdown(markdown);
    const titleBlock = blocks.find(block => block.type === "heading1");
    const title = titleBlock ? titleBlock.text : path.basename(inputPath, path.extname(inputPath));
    const tempDir = packageDir || fs.mkdtempSync(path.join(os.tmpdir(), "md-docx-"));
    const tempZipPath = path.join(os.tmpdir(), `md-docx-${Date.now()}.zip`);

    try {
        if (packageDir) {
            fs.rmSync(tempDir, { recursive: true, force: true });
            ensureDir(tempDir);
        }

        writePackage(tempDir, title, buildDocumentXml(blocks), buildStylesXml());

        if (packageDir) {
            console.log(`DOCX package created: ${tempDir}`);
            return;
        }

        ensureDir(path.dirname(outputPath));
        createZipFromDir(tempDir, tempZipPath);
        fs.copyFileSync(tempZipPath, outputPath);
        console.log(`DOCX created: ${outputPath}`);
    } finally {
        if (!packageDir) {
            fs.rmSync(tempDir, { recursive: true, force: true });
        }
        if (fs.existsSync(tempZipPath)) {
            fs.rmSync(tempZipPath, { force: true });
        }
    }
}

main();
