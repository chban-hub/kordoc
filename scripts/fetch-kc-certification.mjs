#!/usr/bin/env node
/**
 * SafetyKorea KC 안전인증정보 추출 스크립트
 *
 * 공개 웹 검색(기본, API 키 불필요):
 *   node scripts/fetch-kc-certification.mjs --cert-num ZU10040-26007
 *   node scripts/fetch-kc-certification.mjs --model-name XG27JCEG --detail
 *   node scripts/fetch-kc-certification.mjs --list-only --limit 5
 *
 * Open API(인증키 + 명세서 URL 필요):
 *   node scripts/fetch-kc-certification.mjs --mode api \
 *     --env /path/to/.env \
 *     --url "https://www.safetykorea.kr/....json?certNum=ZU10040-26007"
 */

import { readFileSync, existsSync } from "fs"
import { resolve } from "path"

const BASE = "https://www.safetykorea.kr"
const USER_AGENT = "kordoc-kc-cert-fetch/1.0"

function parseArgs(argv) {
  const args = {
    mode: "web",
    env: null,
    url: null,
    certNum: null,
    productName: null,
    modelName: null,
    makerName: null,
    importerName: null,
    pageNo: 0,
    limit: 10,
    listOnly: false,
    detail: false,
    json: false,
  }

  for (let i = 2; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === "--mode") args.mode = argv[++i]
    else if (arg === "--env") args.env = argv[++i]
    else if (arg === "--url") args.url = argv[++i]
    else if (arg === "--cert-num") args.certNum = argv[++i]
    else if (arg === "--product-name") args.productName = argv[++i]
    else if (arg === "--model-name") args.modelName = argv[++i]
    else if (arg === "--maker-name") args.makerName = argv[++i]
    else if (arg === "--importer-name") args.importerName = argv[++i]
    else if (arg === "--page") args.pageNo = Number(argv[++i] ?? 0)
    else if (arg === "--limit") args.limit = Number(argv[++i] ?? 10)
    else if (arg === "--list-only") args.listOnly = true
    else if (arg === "--detail") args.detail = true
    else if (arg === "--json") args.json = true
    else if (arg === "--help" || arg === "-h") {
      printHelp()
      process.exit(0)
    }
  }

  if (args.certNum && !args.detail && !args.listOnly) args.detail = true
  return args
}

function printHelp() {
  console.log(`SafetyKorea KC 안전인증정보 추출

공개 웹 검색 (API 키 불필요):
  node scripts/fetch-kc-certification.mjs --cert-num ZU10040-26007
  node scripts/fetch-kc-certification.mjs --model-name XG27JCEG --detail
  node scripts/fetch-kc-certification.mjs --list-only --limit 5

Open API (인증키 + 명세서 URL):
  node scripts/fetch-kc-certification.mjs --mode api \\
    --env /path/to/.env \\
    --url "https://www.safetykorea.kr/....json?certNum=ZU10040-26007"

옵션:
  --mode web|api     조회 방식 (기본: web)
  --cert-num         KC 인증번호
  --product-name     제품명
  --model-name       모델명
  --maker-name       제조사명
  --importer-name    수입업체명
  --page N           페이지 (기본 0)
  --limit N          최대 건수 (기본 10)
  --list-only        목록만 출력
  --detail           상세 정보까지 조회
  --json             JSON 형식 출력
  --env PATH         .env 파일 (api 모드)
  --url URL          Open API URL (api 모드)
`)
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

function decodeHtml(text) {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ")
    .trim()
}

function stripTags(html) {
  return decodeHtml(html.replace(/<[^>]+>/g, " "))
}

function parseSearchList(html) {
  const rows = []
  const rowRe = /goDetail\('([^']+)',\s*'(\d+)'\)[^]*?<td[^>]*>\s*(\d+)\s*<\/td>[^]*?<a[^>]*>([^<]*)<\/a>[^]*?<td[^>]*>([^<]*)<\/td>[^]*?<td[^>]*>([^<]*)<\/td>[^]*?id="certNum_\d+"[^>]*>\s*([^<\s]+)/g

  for (const match of html.matchAll(rowRe)) {
    rows.push({
      certNum: match[1].trim(),
      certUid: match[2].trim(),
      index: Number(match[3]),
      modelName: stripTags(match[4]),
      productName: stripTags(match[5]),
      certStatus: stripTags(match[6]),
    })
  }

  if (rows.length > 0) return rows

  // 최신 목록 페이지(itemSearch) fallback
  const itemRe = /goDetail\('([^']+)',\s*'(\d+)'\)[^]*?<a[^>]*>([^<]*)<\/a>[^]*?<td[^>]*>([^<]*)<\/td>[^]*?<td[^>]*>([^<]*)<\/td>[^]*?id="certNum_\d+"[^>]*>\s*([^<\s]+)/g
  let i = 1
  for (const match of html.matchAll(itemRe)) {
    rows.push({
      certNum: match[1].trim(),
      certUid: match[2].trim(),
      index: i++,
      modelName: stripTags(match[3]),
      productName: stripTags(match[4]),
      certStatus: stripTags(match[5]),
    })
  }
  return rows
}

function parseDetailTables(html) {
  const sections = {}
  let current = "기타"

  const sectionRe = /<p class="tit">([^<]+)<\/p>|인증정보|제품정보|제조사정보|제조공장|연관 인증 번호/g
  const tableRe = /<table[^>]*>([\s\S]*?)<\/table>/gi

  const sectionMarkers = [...html.matchAll(/<p class="tit">([^<]+)<\/p>/g)].map(m => ({
    title: stripTags(m[1]),
    index: m.index ?? 0,
  }))

  for (const tableMatch of html.matchAll(tableRe)) {
    const tableHtml = tableMatch[1]
    const tableIndex = tableMatch.index ?? 0
    let section = "기타"
    for (const marker of sectionMarkers) {
      if (marker.index <= tableIndex) section = marker.title
      else break
    }

    const pairs = {}
    const rowRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi
    for (const rowMatch of tableHtml.matchAll(rowRe)) {
      const cells = [...rowMatch[1].matchAll(/<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi)].map(m => stripTags(m[1]))
      if (cells.length === 2) {
        pairs[cells[0]] = cells[1]
      } else if (cells.length === 4) {
        pairs[cells[0]] = cells[1]
        pairs[cells[2]] = cells[3]
      }
    }

    if (Object.keys(pairs).length > 0) {
      sections[section] = { ...(sections[section] ?? {}), ...pairs }
    }
  }

  return sections
}

function flattenDetail(sections) {
  const flat = {}
  for (const fields of Object.values(sections)) {
    Object.assign(flat, fields)
  }
  return flat
}

async function fetchText(url, init = {}) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), 20000)
  try {
    const res = await fetch(url, {
      ...init,
      headers: {
        "User-Agent": USER_AGENT,
        ...(init.headers ?? {}),
      },
      signal: controller.signal,
    })
    const text = await res.text()
    return { ok: res.ok, status: res.status, text }
  } finally {
    clearTimeout(timer)
  }
}

async function searchWeb(params) {
  const body = new URLSearchParams()
  if (params.certNum) body.set("certNum", params.certNum.replace(/\s/g, ""))
  if (params.productName) body.set("productName", params.productName.replace(/\s/g, ""))
  if (params.modelName) body.set("modelName", params.modelName)
  if (params.makerName) body.set("makerName", params.makerName.replace(/\s/g, ""))
  if (params.importerName) body.set("importerName", params.importerName.replace(/\s/g, ""))
  body.set("pageNo", String(params.pageNo ?? 0))

  const hasQuery = params.certNum || params.productName || params.modelName || params.makerName || params.importerName

  if (!hasQuery) {
    const { text } = await fetchText(`${BASE}/release/itemSearch`, { method: "GET" })
    return parseSearchList(text)
  }

  const { text } = await fetchText(`${BASE}/release/certificationsearch`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  })
  return parseSearchList(text)
}

async function fetchDetailWeb(certNum) {
  const { text } = await fetchText(`${BASE}/search/searchPop?certNum=${encodeURIComponent(certNum)}`)
  const sections = parseDetailTables(text)
  return {
    certNum,
    sections,
    summary: flattenDetail(sections),
  }
}

async function fetchApi(url, apiKey) {
  const { status, text } = await fetchText(url, {
    headers: { AuthKey: apiKey },
  })

  let json = null
  try {
    json = JSON.parse(text)
  } catch {
    return { status, raw: text.slice(0, 500), error: "JSON 파싱 실패" }
  }

  return {
    status,
    resultCode: json?.resultCode ?? json?.body?.resultCode ?? null,
    resultMsg: json?.resultMsg ?? json?.body?.resultMsg ?? null,
    data: json?.resultData ?? json?.body?.resultData ?? json,
  }
}

function printList(items) {
  console.log(`\n=== KC 인증 목록 (${items.length}건) ===`)
  for (const item of items) {
    console.log(`\n[${item.index}] ${item.certNum} (${item.certStatus})`)
    console.log(`  제품명: ${item.productName}`)
    console.log(`  모델명: ${item.modelName}`)
    if (item.certUid) console.log(`  UID: ${item.certUid}`)
  }
}

function printDetail(detail) {
  console.log(`\n=== KC 인증 상세: ${detail.certNum} ===`)
  for (const [section, fields] of Object.entries(detail.sections)) {
    console.log(`\n[${section}]`)
    for (const [key, value] of Object.entries(fields)) {
      if (value) console.log(`  ${key}: ${value}`)
    }
  }
}

async function runWeb(args) {
  const items = await searchWeb(args)
  const limited = items.slice(0, args.limit)

  if (limited.length === 0) {
    console.log("검색 결과가 없습니다.")
    return { mode: "web", count: 0, items: [], details: [] }
  }

  let details = []
  if (args.detail && !args.listOnly) {
    const targets = args.certNum
      ? [{ certNum: args.certNum, certUid: limited[0]?.certUid }]
      : limited

    for (const item of targets) {
      details.push(await fetchDetailWeb(item.certNum))
    }
  }

  const result = { mode: "web", count: limited.length, items: limited, details }

  if (args.json) {
    console.log(JSON.stringify(result, null, 2))
    return result
  }

  printList(limited)
  for (const detail of details) printDetail(detail)

  console.log("\n=== 실현 가능성 ===")
  console.log("✅ 공개 웹 검색/상세 조회로 KC 안전인증정보 추출이 가능합니다.")
  console.log("   - 목록: POST /release/certificationsearch")
  console.log("   - 상세: GET  /search/searchPop?certNum=...")
  console.log("   - API 키 없이 인증번호·제품명·모델명 검색 지원")

  return result
}

async function runApi(args) {
  if (!args.env || !existsSync(args.env)) {
    throw new Error("--env 경로가 필요합니다.")
  }
  if (!args.url) {
    throw new Error("--url 로 Open API 엔드포인트를 지정하세요. (사용설명서 3.2절)")
  }

  const env = loadEnv(resolve(args.env))
  const keyInfo = findApiKey(env)
  if (!keyInfo?.value) {
    throw new Error(".env에서 Safetykorea_APi_key를 찾지 못했습니다.")
  }

  const apiResult = await fetchApi(args.url, keyInfo.value)
  const result = { mode: "api", ...apiResult }

  if (args.json) {
    console.log(JSON.stringify(result, null, 2))
    return result
  }

  console.log("=== Open API 조회 ===")
  console.log(`HTTP ${result.status}`)
  console.log(`resultCode: ${result.resultCode ?? "(없음)"}`)
  console.log(`resultMsg: ${result.resultMsg ?? "(없음)"}`)
  if (result.error) console.log(`error: ${result.error}`)
  if (result.data) {
    console.log("\nresultData:")
    console.log(JSON.stringify(result.data, null, 2).slice(0, 2000))
  }

  const code = String(result.resultCode ?? "")
  if (code === "2000") {
    console.log("\n✅ Open API로 안전인증정보 추출이 가능합니다.")
  } else if (code === "4000") {
    console.log("\n❌ 인증키가 유효하지 않습니다.")
  } else if (code === "4001") {
    console.log("\n⚠️  인증키는 인식되지만 IP가 허용되지 않습니다.")
  } else if (code === "2004" || code === "4005") {
    console.log("\n⚠️  API 연결은 되지만 조회 파라미터를 확인하세요.")
  } else {
    console.log("\n❓ API URL이 명세서와 일치하는지 확인하세요.")
  }

  return result
}

async function main() {
  const args = parseArgs(process.argv)

  if (args.mode === "api") {
    await runApi(args)
    return
  }

  await runWeb(args)
}

main().catch(err => {
  console.error("FATAL:", err instanceof Error ? err.message : err)
  process.exit(1)
})
