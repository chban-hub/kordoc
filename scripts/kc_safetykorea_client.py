"""SafetyKorea KC 인증정보 조회 클라이언트 (searchPop 상세 API)."""

from __future__ import annotations

import html
import re
import ssl
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Optional

BASE_URL = "https://www.safetykorea.kr"
SEARCH_POP_PATH = "/search/searchPop"
USER_AGENT = "kordoc-kc-cert-validate/1.0 (python)"


@dataclass
class KCCertDetail:
    cert_num: str
    cert_status: str
    product_name: str
    model_name: str
    recall_status: str
    cert_date: str
    cert_org: str
    raw_fields: dict[str, str]
    found: bool


def strip_tags(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def create_ssl_context(*, insecure: bool = False) -> ssl.SSLContext:
    if insecure:
        return ssl._create_unverified_context()
    return ssl.create_default_context()


def fetch_html(
    url: str,
    *,
    ssl_context: Optional[ssl.SSLContext] = None,
) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT}, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=20, context=ssl_context) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"HTTP {e.code}: {url}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"요청 실패: {e.reason}") from e


def parse_detail_tables(page_html: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    section_markers = [
        (m.start(), strip_tags(m.group(1)))
        for m in re.finditer(r'<p class="tit">([^<]+)</p>', page_html)
    ]

    for table_match in re.finditer(r"<table[^>]*>([\s\S]*?)</table>", page_html, re.IGNORECASE):
        table_html = table_match.group(1)
        table_index = table_match.start()
        section = "기타"
        for marker_index, title in section_markers:
            if marker_index <= table_index:
                section = title

        for row_match in re.finditer(r"<tr[^>]*>([\s\S]*?)</tr>", table_html, re.IGNORECASE):
            cells = [
                strip_tags(m.group(1))
                for m in re.finditer(r"<t[hd][^>]*>([\s\S]*?)</t[hd]>", row_match.group(1), re.IGNORECASE)
            ]
            if len(cells) == 2:
                key = f"{section}.{cells[0]}" if section != "기타" else cells[0]
                fields[key] = cells[1]
            elif len(cells) == 4:
                fields[cells[0]] = cells[1]
                fields[cells[2]] = cells[3]

    return fields


def _pick_field(fields: dict[str, str], *names: str) -> str:
    for name in names:
        if name in fields and fields[name]:
            return fields[name]
        for key, value in fields.items():
            if key.endswith(f".{name}") and value:
                return value
    return ""


def fetch_cert_detail(cert_num: str, *, insecure: bool = False) -> KCCertDetail:
    """searchPop API로 KC 인증 상세 조회."""
    cert_num = cert_num.strip()
    url = f"{BASE_URL}{SEARCH_POP_PATH}?certNum={urllib.parse.quote(cert_num)}"
    ssl_context = create_ssl_context(insecure=insecure)
    page_html = fetch_html(url, ssl_context=ssl_context)

    if "인증정보가 존재하지 않습니다" in page_html or "noDataWrap" in page_html:
        return KCCertDetail(
            cert_num=cert_num,
            cert_status="",
            product_name="",
            model_name="",
            recall_status="",
            cert_date="",
            cert_org="",
            raw_fields={},
            found=False,
        )

    fields = parse_detail_tables(page_html)
    recall = _pick_field(fields, "리콜현황 (모델명)", "리콜현황", "리콜현황<br/>(모델명)")

    return KCCertDetail(
        cert_num=_pick_field(fields, "인증번호") or cert_num,
        cert_status=_pick_field(fields, "인증상태", "인증현황"),
        product_name=_pick_field(fields, "제품명"),
        model_name=_pick_field(fields, "모델명"),
        recall_status=recall,
        cert_date=_pick_field(fields, "인증일자"),
        cert_org=_pick_field(fields, "인증기관"),
        raw_fields=fields,
        found=True,
    )
