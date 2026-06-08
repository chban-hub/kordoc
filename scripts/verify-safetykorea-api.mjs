#!/usr/bin/env node
/**
 * SafetyKorea Open API 인증키 검증 스크립트
 *
 * 사용법:
 *   node scripts/verify-safetykorea-api.mjs \
 *     --env /Users/chban/Environment/.env \
 *     --docx "/Users/chban/Downloads/Open_API_사용설명서_v2.0.docx"
 *
 * 선택:
 *   --url "https://www.safetykorea.kr/....json?pageIndex=1&pageSize=1"
 */

import { readFileSync, existsSync } from "fs"
import { resolve } from "path"
import { parse } from "../dist/index.js"

const RESULT_MESSAGES = {
  2000: "Success — 인증키 정상",
  2004: "No Date — 조회 결과 없음 (키는 유효할 수 있음)",
  4000: "Invalid Auth Key — 인증키 오류/만료",
  4001: "Invalid IP — 인증키는 맞지만 허용 IP가 아님",
  4005: "Invalid Parameter — 파라미터 오류 (키는 유효할 수 있음)",
  5000: "Internal Server Error — 서버 오류",
}

function parseArgs(argv) {
  const args = { env: null, docx: null, urls: [] }
  for (let i = 2; i < argv.length; i++) {
    if (argv[i] === "--env") args.env = argv[++i]
    else if (argv[i] === "--docx") args.docx = argv[++i]
    else if (argv[i] === "--url") args.urls.push(argv[++i])
  }
  return args
}

function loadEnv(path) {
  const content = readFileSync(path, "utf8")
  const vars = {}
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim()
    if (!trimmed || trimmed.startsWith("#")) continue
    const idx = trimmed.indexOf("=")
    if (idx < 0) continue
    const key = trimmed.slice(0, idx).trim()
    let value = trimmed.slice(idx + 1).trim()
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1)
    }
    vars[key] = value
  }
  return vars
}

function findApiKey(env) {
  const entries = Object.entries(env)
  const preferred = entries.find(([k]) => /^Safetykorea_APi_key$/i.test(k))
  if (preferred?.[1]) return { name: preferred[0], value: preferred[1] }

  const fuzzy = entries.find(([k, v]) => /safetykorea/i.test(k) && /api|auth|key/i.test(k) && v)
  if (fuzzy) return { name: fuzzy[0], value: fuzzy[1] }

  return null
}

function extractApiUrls(text) {
  const urls = new Set()
  const re = /https?:\/\/(?:www\.)?safetykorea\.kr[^\s"'<>|]+/gi
  for (const match of text.matchAll(re)) {
    let url = match[0].replace(/[.,;]+$/, "")
    if (/openApi|\.json|\.xml/i.test(url) && !/certificationsearch|searchPop|subPage/i.test(url)) {
      urls.add(url)
    }
  }
  return [...urls]
}

async function extractUrlsFromDocx(docxPath) {
  const result = await parse(docxPath)
  if (!result.success) {
    throw new Error(`DOCX 파싱 실패: ${result.error}`)
  }
  const blob = `${result.markdown}\n${JSON.stringify(result.blocks)}`
  return extractApiUrls(blob)
}

async function callApi(url, apiKey) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), 15000)
  try {
    const res = await fetch(url, {
      method: "GET",
      headers: { AuthKey: apiKey },
      signal: controller.signal,
    })
    const text = await res.text()
    let json = null
    try { json = JSON.parse(text) } catch { /* xml or html */ }

    const resultCode = json?.resultCode ?? json?.body?.resultCode
    const resultMsg = json?.resultMsg ?? json?.body?.resultMsg

    return {
      url,
      httpStatus: res.status,
      resultCode: resultCode != null ? String(resultCode) : null,
      resultMsg: resultMsg ?? null,
      ok: resultCode === 2000 || resultCode === "2000",
      maybeValidKey: ["2000", "2004", "4005"].includes(String(resultCode)),
      invalidKey: resultCode === 4000 || resultCode === "4000",
      invalidIp: resultCode === 4001 || resultCode === "4001",
      preview: text.slice(0, 400),
    }
  } finally {
    clearTimeout(timer)
  }
}

function printResult(result) {
  console.log(`\n[요청] ${result.url}`)
  console.log(`  HTTP ${result.httpStatus}`)
  if (result.resultCode != null) {
    const hint = RESULT_MESSAGES[Number(result.resultCode)] ?? RESULT_MESSAGES[result.resultCode] ?? ""
    console.log(`  resultCode: ${result.resultCode}${hint ? ` (${hint})` : ""}`)
  }
  if (result.resultMsg) console.log(`  resultMsg: ${result.resultMsg}`)
  if (!result.resultCode) {
    console.log("  응답 미리보기:")
    console.log("  " + result.preview.replace(/\n/g, "\n  "))
  }
}

async function main() {
  const args = parseArgs(process.argv)

  if (!args.env || !existsSync(args.env)) {
    console.error("ERROR: --env 경로가 필요합니다. 예: --env /Users/chban/Environment/.env")
    process.exit(1)
  }

  const env = loadEnv(resolve(args.env))
  const keyInfo = findApiKey(env)
  if (!keyInfo?.value) {
    console.error("ERROR: .env에서 SafetyKorea API 키를 찾지 못했습니다.")
    console.error("       예상 변수명: Safetykorea_APi_key")
    process.exit(1)
  }

  console.log("=== SafetyKorea Open API 키 검증 ===")
  console.log(`환경변수: ${keyInfo.name}`)
  console.log(`키 길이: ${keyInfo.value.length}자 (값은 출력하지 않음)`)

  let urls = [...args.urls]

  if (args.docx && existsSync(args.docx)) {
    console.log(`\nDOCX에서 API URL 추출 중: ${args.docx}`)
    const found = await extractUrlsFromDocx(resolve(args.docx))
    console.log(`  추출된 API URL: ${found.length}개`)
    urls.push(...found)
  }

  urls = [...new Set(urls)]

  if (urls.length === 0) {
    console.error("\nERROR: 테스트할 API URL이 없습니다.")
    console.error("  1) --docx 로 사용설명서를 지정하거나")
    console.error("  2) --url 로 명세서의 KC인증/리콜 조회 URL을 직접 지정하세요.")
    console.error("\n명세서 v2.0 인증 방식:")
    console.error("  - HTTP Header: AuthKey: <발급받은 서비스 ID>")
    console.error("  - 성공 코드: resultCode=2000")
    console.error("  - 키 오류: resultCode=4000, IP 오류: resultCode=4001")
    process.exit(1)
  }

  const results = []
  for (const url of urls) {
    const result = await callApi(url, keyInfo.value)
    printResult(result)
    results.push(result)
  }

  const anySuccess = results.some(r => r.ok)
  const anyMaybeValid = results.some(r => r.maybeValidKey)
  const allInvalidKey = results.length > 0 && results.every(r => r.invalidKey)
  const anyInvalidIp = results.some(r => r.invalidIp)

  console.log("\n=== 종합 판정 ===")
  if (anySuccess) {
    console.log("✅ API 키가 정상 작동합니다 (resultCode 2000).")
    process.exit(0)
  }
  if (anyInvalidIp) {
    console.log("⚠️  인증키는 인식되지만 IP가 허용되지 않습니다 (4001).")
    console.log("   신청 시 등록한 서버 IP와 현재 요청 IP를 확인하세요.")
    process.exit(2)
  }
  if (anyMaybeValid) {
    console.log("⚠️  인증키는 유효해 보이지만 조회 조건/파라미터를 확인하세요 (2004 또는 4005).")
    process.exit(0)
  }
  if (allInvalidKey) {
    console.log("❌ 인증키가 유효하지 않습니다 (4000 Invalid Auth Key).")
    process.exit(1)
  }
  console.log("❓ API 응답을 해석할 수 없습니다. URL이 명세서와 일치하는지 확인하세요.")
  process.exit(3)
}

main().catch(err => {
  console.error("FATAL:", err instanceof Error ? err.message : err)
  process.exit(1)
})
