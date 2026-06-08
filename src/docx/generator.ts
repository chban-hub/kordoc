/**
 * IRBlock[] → DOCX 역변환
 *
 * 지원: 단락, 헤딩, 테이블, 리스트, 구분선, 이미지(기본)
 * jszip으로 Office Open XML 패키징.
 */

import JSZip from "jszip"
import type { ExtractedImage, IRBlock, IRTable, IRCell } from "../types.js"

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

export interface DocxGeneratorOptions {
  title?: string
  /** parse() 결과의 images 배열 — image_N.png 참조와 매칭 */
  images?: ExtractedImage[]
}

/**
 * IRBlock 배열을 DOCX (ArrayBuffer)로 변환.
 */
export async function blocksToDocx(blocks: IRBlock[], options?: DocxGeneratorOptions): Promise<ArrayBuffer> {
  const imageMap = new Map<string, ExtractedImage>()
  for (const img of options?.images ?? []) {
    imageMap.set(img.filename, img)
  }

  const mediaFiles: { filename: string; data: Uint8Array; contentType: string }[] = []
  const relationships: { id: string; type: string; target: string }[] = []
  let nextRelId = 2

  const bodyParts: string[] = []

  for (const block of blocks) {
    switch (block.type) {
      case "heading":
        if (block.text?.trim()) {
          const level = Math.min(Math.max(block.level ?? 1, 1), 6)
          bodyParts.push(generateHeading(block.text, level))
        }
        break

      case "paragraph":
        if (block.text?.trim()) {
          bodyParts.push(generateParagraph(block.text, block.style, block.href, block.footnoteText))
        }
        break

      case "list":
        if (block.text?.trim()) {
          bodyParts.push(generateListItem(block.text, block.listType === "ordered"))
          if (block.children) {
            for (const child of block.children) {
              if (child.text?.trim()) {
                bodyParts.push(generateListItem(child.text, child.listType === "ordered", 1))
              }
            }
          }
        }
        break

      case "separator":
        bodyParts.push(generateSeparator())
        break

      case "table":
        if (block.table) {
          bodyParts.push(generateTable(block.table))
        }
        break

      case "image": {
        const ref = block.text?.trim()
        if (!ref) break
        const extracted = imageMap.get(ref)
        const imgData = extracted?.data ?? block.imageData?.data
        const imgMime = extracted?.mimeType ?? block.imageData?.mimeType
        if (!imgData || !imgMime) {
          bodyParts.push(generateParagraph(`[이미지: ${ref}]`))
          break
        }
        const ext = mimeToExt(imgMime)
        const mediaName = `image${mediaFiles.length + 1}.${ext}`
        mediaFiles.push({
          filename: mediaName,
          data: imgData,
          contentType: imgMime,
        })
        const relId = `rId${nextRelId++}`
        relationships.push({
          id: relId,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          target: `media/${mediaName}`,
        })
        bodyParts.push(generateImage(relId, mediaName))
        break
      }
    }
  }

  if (bodyParts.length === 0) {
    bodyParts.push(generateParagraph(""))
  }

  const zip = new JSZip()
  const hasMedia = mediaFiles.length > 0

  zip.file("[Content_Types].xml", generateContentTypes(hasMedia))
  zip.file("_rels/.rels", generateRootRels())
  zip.file("word/document.xml", generateDocument(bodyParts.join("\n")))
  zip.file("word/_rels/document.xml.rels", generateDocumentRels(relationships))
  zip.file("word/styles.xml", generateStyles())
  zip.file("word/numbering.xml", generateNumbering())
  zip.file("docProps/core.xml", generateCoreProps(options?.title))
  zip.file("docProps/app.xml", generateAppProps())

  for (const media of mediaFiles) {
    zip.file(`word/media/${media.filename}`, media.data)
  }

  return await zip.generateAsync({ type: "arraybuffer" })
}

// ─── XML 생성 ────────────────────────────────────────

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
}

function escapeXmlText(text: string): string {
  // Word XML: 공백 보존이 필요한 경우 xml:space="preserve"
  const escaped = escapeXml(text)
  if (/^\s|\s$/.test(text) || text.includes("  ")) {
    return `<w:t xml:space="preserve">${escaped}</w:t>`
  }
  return `<w:t>${escaped}</w:t>`
}

function generateRun(text: string, style?: { bold?: boolean; italic?: boolean }): string {
  const rPr: string[] = []
  if (style?.bold) rPr.push("<w:b/>")
  if (style?.italic) rPr.push("<w:i/>")
  const rPrXml = rPr.length > 0 ? `<w:rPr>${rPr.join("")}</w:rPr>` : ""
  return `<w:r>${rPrXml}${escapeXmlText(text)}</w:r>`
}

function generateParagraph(
  text: string,
  style?: { bold?: boolean; italic?: boolean },
  href?: string,
  footnoteText?: string,
): string {
  let content = text
  if (footnoteText) content += ` (주: ${footnoteText})`

  if (href) {
    const runs = [
      generateRun(content, style),
      generateRun(` (${href})`, { italic: true }),
    ]
    return `<w:p>${runs.join("")}</w:p>`
  }

  // 줄바꿈 처리
  const lines = content.split("\n")
  if (lines.length === 1) {
    return `<w:p>${generateRun(content, style)}</w:p>`
  }

  const runs: string[] = []
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) runs.push("<w:br/>")
    if (lines[i]) runs.push(generateRun(lines[i], style))
  }
  return `<w:p>${runs.join("")}</w:p>`
}

function generateHeading(text: string, level: number): string {
  return `<w:p>
    <w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>
    ${generateRun(text, { bold: true })}
  </w:p>`
}

function generateListItem(text: string, ordered: boolean, indent = 0): string {
  const numId = ordered ? 1 : 2
  const ilvl = indent
  return `<w:p>
    <w:pPr>
      <w:numPr>
        <w:ilvl w:val="${ilvl}"/>
        <w:numId w:val="${numId}"/>
      </w:numPr>
      <w:ind w:left="${720 * (indent + 1)}" w:hanging="360"/>
    </w:pPr>
    ${generateRun(text)}
  </w:p>`
}

function generateSeparator(): string {
  return `<w:p>
    <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr>
  </w:p>`
}

function generateTable(table: IRTable): string {
  const { rows, cols, cells } = table
  if (rows === 0 || cols === 0) return ""

  const gridCols = Array.from({ length: cols }, () =>
    `<w:gridCol w:w="${Math.floor(9000 / cols)}"/>`
  ).join("")

  const trElements: string[] = []
  const colSkip = new Set<string>()

  for (let r = 0; r < rows; r++) {
    const tcElements: string[] = []

    for (let c = 0; c < cols; c++) {
      if (colSkip.has(`${r},${c}`)) continue

      const cell = cells[r]?.[c]
      if (!cell) continue

      const tcPr: string[] = []
      let isContinue = false

      for (let up = r - 1; up >= 0; up--) {
        const above = cells[up]?.[c]
        if (above && above.rowSpan > 1 && up + above.rowSpan > r) {
          tcPr.push(`<w:vMerge w:val="continue"/>`)
          isContinue = true
          break
        }
      }

      if (!isContinue) {
        if (cell.colSpan > 1) tcPr.push(`<w:gridSpan w:val="${cell.colSpan}"/>`)
        if (cell.rowSpan > 1) {
          tcPr.push(`<w:vMerge w:val="restart"/>`)
          for (let dc = 1; dc < cell.colSpan; dc++) {
            colSkip.add(`${r},${c + dc}`)
          }
        }
      }

      const cellText = isContinue ? "" : cell.text.trim()
      const pContent = cellText
        ? cellText.split("\n").map(line => generateParagraph(line)).join("")
        : generateParagraph("")

      const tcPrXml = tcPr.length > 0 ? `<w:tcPr>${tcPr.join("")}</w:tcPr>` : ""
      tcElements.push(`<w:tc>${tcPrXml}${pContent}</w:tc>`)
    }

    if (tcElements.length > 0) {
      trElements.push(`<w:tr>${tcElements.join("")}</w:tr>`)
    }
  }

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="0" w:type="auto"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid>${gridCols}</w:tblGrid>
    ${trElements.join("\n")}
  </w:tbl>`
}

function generateImage(relId: string, _mediaName: string): string {
  const cx = 4000000
  const cy = 3000000
  return `<w:p>
    <w:r>
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${WP_NS}">
          <wp:extent cx="${cx}" cy="${cy}"/>
          <wp:docPr id="1" name="Picture"/>
          <a:graphic xmlns:a="${A_NS}">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic xmlns:pic="${PIC_NS}">
                <pic:nvPicPr>
                  <pic:cNvPr id="0" name="Picture"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip xmlns:r="${R_NS}" r:embed="${relId}"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="${cx}" cy="${cy}"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>`
}

function generateContentTypes(hasMedia: boolean): string {
  const mediaTypes = hasMedia
    ? `
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="bmp" ContentType="image/bmp"/>`
    : ""

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>${mediaTypes}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
}

function generateRootRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
}

function generateDocumentRels(imageRels: { id: string; type: string; target: string }[]): string {
  const base = [
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    `<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`,
  ]
  const images = imageRels.map(rel =>
    `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"/>`
  )
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${[...base, ...images].join("\n  ")}
</Relationships>`
}

function generateDocument(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">
  <w:body>
    ${body}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`
}

function generateStyles(): string {
  const headings = [1, 2, 3, 4, 5, 6].map(n => `
  <w:style w:type="paragraph" w:styleId="Heading${n}">
    <w:name w:val="heading ${n}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr><w:outlineLvl w:val="${n - 1}"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="${36 - n * 2}"/></w:rPr>
  </w:style>`).join("")

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕" w:eastAsia="맑은 고딕"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>${headings}
</w:styles>`
}

function generateNumbering(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`
}

function generateCoreProps(title?: string): string {
  const now = new Date().toISOString()
  const titleXml = title ? `<dc:title>${escapeXml(title)}</dc:title>` : ""
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  ${titleXml}
  <dc:creator>kordoc</dc:creator>
  <cp:lastModifiedBy>kordoc</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`
}

function generateAppProps(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>kordoc</Application>
</Properties>`
}

function mimeToExt(mime: string): string {
  switch (mime) {
    case "image/png": return "png"
    case "image/jpeg": return "jpg"
    case "image/gif": return "gif"
    case "image/bmp": return "bmp"
    case "image/wmf": return "wmf"
    case "image/emf": return "emf"
    default: return "png"
  }
}
