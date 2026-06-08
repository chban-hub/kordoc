#!/usr/bin/env python3
"""
SafetyKorea KC 안전인증정보 목록 추출 (Python)

사용법:
  python scripts/fetch_kc_certification.py --limit 10
  python scripts/fetch_kc_certification.py --cert-num ZU10040-26007
  python scripts/fetch_kc_certification.py --product-name 모니터 --limit 5
  python scripts/fetch_kc_certification.py --limit 10 --json
  python scripts/fetch_kc_certification.py --limit 10 --csv output.csv
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from typing import Optional

BASE_URL = "https://www.safetykorea.kr"
USER_AGENT = "kordoc-kc-cert-fetch/1.0 (python)"


@dataclass
class KCCertItem:
    index: int
    cert_num: str
    cert_uid: str
    model_name: str
    product_name: str
    cert_status: str


def strip_tags(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def fetch_html(url: str, *, method: str = "GET", data: Optional[dict] = None) -> str:
    headers = {"User-Agent": USER_AGENT}
    body = None
    if data is not None:
        body = urllib.parse.urlencode(data).encode("utf-8")
        headers["Content-Type"] = "application/x-www-form-urlencoded"

    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"HTTP {e.code}: {url}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"요청 실패: {e.reason}") from e


def parse_search_list(page_html: str) -> list[KCCertItem]:
    """검색 결과 또는 최신 목록 HTML에서 인증 항목 파싱."""
    items: list[KCCertItem] = []

    # 검색 결과 (certificationsearch)
    search_pattern = re.compile(
        r"goDetail\('([^']+)',\s*'(\d+)'\)"
        r"[\s\S]*?<td[^>]*>\s*(\d+)\s*</td>"
        r"[\s\S]*?<a[^>]*>([^<]*)</a>"
        r"[\s\S]*?<td[^>]*>([^<]*)</td>"
        r"[\s\S]*?<td[^>]*>([^<]*)</td>"
        r"[\s\S]*?id=\"certNum_\d+\"[^>]*>\s*([^<\s]+)",
        re.IGNORECASE,
    )
    for match in search_pattern.finditer(page_html):
        items.append(
            KCCertItem(
                index=int(match.group(3)),
                cert_num=match.group(1).strip(),
                cert_uid=match.group(2).strip(),
                model_name=strip_tags(match.group(4)),
                product_name=strip_tags(match.group(5)),
                cert_status=strip_tags(match.group(6)),
            )
        )

    if items:
        return items

    # 최신 목록 (itemSearch) fallback
    item_pattern = re.compile(
        r"goDetail\('([^']+)',\s*'(\d+)'\)"
        r"[\s\S]*?<a[^>]*>([^<]*)</a>"
        r"[\s\S]*?<td[^>]*>([^<]*)</td>"
        r"[\s\S]*?<td[^>]*>([^<]*)</td>"
        r"[\s\S]*?id=\"certNum_\d+\"[^>]*>\s*([^<\s]+)",
        re.IGNORECASE,
    )
    for i, match in enumerate(item_pattern.finditer(page_html), start=1):
        items.append(
            KCCertItem(
                index=i,
                cert_num=match.group(1).strip(),
                cert_uid=match.group(2).strip(),
                model_name=strip_tags(match.group(3)),
                product_name=strip_tags(match.group(4)),
                cert_status=strip_tags(match.group(5)),
            )
        )

    return items


def fetch_latest_list(limit: int = 10) -> list[KCCertItem]:
    """최신 KC 인증 목록 (API 키 불필요)."""
    html_text = fetch_html(f"{BASE_URL}/release/itemSearch")
    return parse_search_list(html_text)[:limit]


def search_certifications(
    *,
    cert_num: Optional[str] = None,
    product_name: Optional[str] = None,
    model_name: Optional[str] = None,
    maker_name: Optional[str] = None,
    importer_name: Optional[str] = None,
    page_no: int = 0,
    limit: int = 10,
) -> list[KCCertItem]:
    """조건별 KC 인증 검색."""
    has_query = any([cert_num, product_name, model_name, maker_name, importer_name])
    if not has_query:
        return fetch_latest_list(limit)

    data: dict[str, str] = {"pageNo": str(page_no)}
    if cert_num:
        data["certNum"] = cert_num.replace(" ", "")
    if product_name:
        data["productName"] = product_name.replace(" ", "")
    if model_name:
        data["modelName"] = model_name
    if maker_name:
        data["makerName"] = maker_name.replace(" ", "")
    if importer_name:
        data["importerName"] = importer_name.replace(" ", "")

    html_text = fetch_html(
        f"{BASE_URL}/release/certificationsearch",
        method="POST",
        data=data,
    )
    return parse_search_list(html_text)[:limit]


def print_list(items: list[KCCertItem]) -> None:
    print(f"\n=== KC 인증 목록 ({len(items)}건) ===\n")
    for item in items:
        print(f"[{item.index}] {item.cert_num} ({item.cert_status})")
        print(f"  제품명: {item.product_name}")
        print(f"  모델명: {item.model_name}")
        if item.cert_uid:
            print(f"  UID: {item.cert_uid}")
        print()


def write_csv(items: list[KCCertItem], path: str) -> None:
    fieldnames = ["index", "cert_num", "cert_uid", "product_name", "model_name", "cert_status"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for item in items:
            writer.writerow(asdict(item))
    print(f"CSV 저장: {path} ({len(items)}건)")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="SafetyKorea KC 안전인증정보 목록 추출 (API 키 불필요)",
    )
    parser.add_argument("--cert-num", help="KC 인증번호")
    parser.add_argument("--product-name", help="제품명")
    parser.add_argument("--model-name", help="모델명")
    parser.add_argument("--maker-name", help="제조사명")
    parser.add_argument("--importer-name", help="수입업체명")
    parser.add_argument("--page", type=int, default=0, help="페이지 번호 (기본 0)")
    parser.add_argument("--limit", type=int, default=10, help="최대 건수 (기본 10)")
    parser.add_argument("--json", action="store_true", help="JSON 출력")
    parser.add_argument("--csv", metavar="FILE", help="CSV 파일로 저장")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        items = search_certifications(
            cert_num=args.cert_num,
            product_name=args.product_name,
            model_name=args.model_name,
            maker_name=args.maker_name,
            importer_name=args.importer_name,
            page_no=args.page,
            limit=args.limit,
        )
    except RuntimeError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    if not items:
        print("검색 결과가 없습니다.")
        return 0

    if args.json:
        print(json.dumps([asdict(i) for i in items], ensure_ascii=False, indent=2))
    else:
        print_list(items)

    if args.csv:
        write_csv(items, args.csv)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
