/** DOCX 역변환 (generator) 테스트 — 라운드트립 검증 */

import { describe, it } from "node:test"
import assert from "node:assert/strict"
import { blocksToDocx } from "../src/docx/generator.js"
import { parse } from "../src/index.js"
import type { IRBlock } from "../src/types.js"

describe("blocksToDocx", () => {
  it("단순 단락 → DOCX → 라운드트립", async () => {
    const blocks: IRBlock[] = [
      { type: "paragraph", text: "Open API 사용설명서" },
      { type: "paragraph", text: "본 문서는 테스트 목적의 더미 문서입니다." },
    ]
    const docxBuf = await blocksToDocx(blocks, { title: "테스트 문서" })

    assert.ok(docxBuf instanceof ArrayBuffer)
    assert.ok(docxBuf.byteLength > 0)

    const result = await parse(docxBuf)
    assert.equal(result.success, true, `파싱 실패: ${result.success === false ? result.error : ""}`)
    if (result.success) {
      assert.ok(result.markdown.includes("Open API 사용설명서"))
      assert.ok(result.markdown.includes("더미 문서"))
    }
  })

  it("헤딩 + 테이블 → 라운드트립", async () => {
    const blocks: IRBlock[] = [
      { type: "heading", text: "1. 개요", level: 1 },
      { type: "paragraph", text: "Open API는 외부 시스템에서 서비스에 접근할 수 있도록 합니다." },
      {
        type: "table",
        table: {
          rows: 2,
          cols: 2,
          hasHeader: true,
          cells: [
            [{ text: "항목", colSpan: 1, rowSpan: 1 }, { text: "설명", colSpan: 1, rowSpan: 1 }],
            [{ text: "인증", colSpan: 1, rowSpan: 1 }, { text: "API Key", colSpan: 1, rowSpan: 1 }],
          ],
        },
      },
    ]
    const docxBuf = await blocksToDocx(blocks)
    const result = await parse(docxBuf)

    assert.equal(result.success, true)
    if (result.success) {
      assert.ok(result.markdown.includes("개요"))
      assert.ok(result.markdown.includes("인증"))
      assert.ok(result.blocks.some(b => b.type === "table"))
    }
  })

  it("리스트 블록 생성", async () => {
    const blocks: IRBlock[] = [
      { type: "list", text: "HTTPS POST 방식 사용", listType: "unordered" },
      { type: "list", text: "JSON 형식 사용", listType: "unordered" },
    ]
    const docxBuf = await blocksToDocx(blocks)
    const result = await parse(docxBuf)

    assert.equal(result.success, true)
    if (result.success) {
      assert.ok(result.markdown.includes("HTTPS POST"))
      assert.ok(result.markdown.includes("JSON"))
    }
  })
})
