#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI DX 오프라인 심화캠프 — DX과제 산출물 종합 도구

아웃룩에서 받은 산출물(짚파일·단일파일)을 한 폴더에 모아두고 실행하면
① 짚파일을 한글 파일명 깨짐 없이 풀고
② 제출자별로 귀속시킨 뒤
③ 산출물 종합 HTML(SECO 표준 서식)을 생성한다.

사용법 (Windows PowerShell):
    python camp_deliverables.py --input "C:\\SECO_HR\\...\\산출물"

옵션:
    --input       산출물 원본 폴더 (zip 및 개별 파일이 들어있는 곳)  [필수]
    --extract-dir 압축 해제 위치 (기본: <input>/_해제)
    --output      생성할 HTML 경로 (기본: <input>/산출물_종합.html)
    --roster      대상자 명부 JSON (기본: ../data/심화캠프_대상자_23명.json)
    --docno       문서번호 (기본: HR-2026-DX-교육-11)

표준 라이브러리만 사용한다 — pip 설치 불필요.
"""

import argparse
import html
import json
import os
import sys
import zipfile
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────
# 1. 한글 zip 파일명 복원
# ─────────────────────────────────────────────────────────────
# 윈도우 알집/탐색기에서 만든 zip은 UTF-8 플래그(0x800)가 없으면
# Python zipfile이 파일명을 cp437로 디코딩해 한글이 깨진다.
# cp437로 되돌린 뒤 cp949(euc-kr 확장)로 다시 읽어야 원래 한글이 나온다.

def restore_korean_name(info: zipfile.ZipInfo) -> str:
    """ZipInfo의 파일명을 올바른 한글로 복원한다."""
    name = info.filename
    if info.flag_bits & 0x800:
        return name  # 이미 UTF-8로 기록됨
    for enc in ("cp949", "euc-kr", "utf-8"):
        try:
            fixed = name.encode("cp437").decode(enc)
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        # 복원 결과에 한글이 있으면 성공으로 본다
        if any("\uac00" <= ch <= "\ud7a3" for ch in fixed):
            return fixed
    return name


def safe_join(base: Path, member: str) -> Path:
    """zip slip 방지 — 압축 해제 경로가 base 밖으로 나가지 않게 한다."""
    member = member.replace("\\", "/")
    target = (base / member).resolve()
    if not str(target).startswith(str(base.resolve())):
        raise ValueError(f"압축 경로 이탈 감지: {member}")
    return target


def extract_zip(zip_path: Path, dest_root: Path) -> list[dict]:
    """zip을 풀고 내부 파일 목록을 반환한다."""
    dest = dest_root / zip_path.stem
    dest.mkdir(parents=True, exist_ok=True)
    files = []
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            korean = restore_korean_name(info)
            try:
                target = safe_join(dest, korean)
            except ValueError as e:
                print(f"  ! 건너뜀 — {e}")
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                out.write(src.read())
            files.append({
                "이름": Path(korean).name,
                "경로": korean,
                "크기": info.file_size,
                "실제경로": str(target),
            })
    return files


# ─────────────────────────────────────────────────────────────
# 2. 제출자 귀속
# ─────────────────────────────────────────────────────────────

COMPANY_HINTS = {
    "좋은사람들": ["좋은사람들", "gpin", "GP"],
    "인터렉스메가라인": ["인터렉스", "메가라인", "megaline"],
    "휴맥스해운항공": ["휴맥스", "humex"],
    "인베스터유나이티드": ["인베스터", "유나이티드", "investorunited", "pinewood"],
    "우리인터텍스": ["우리인터텍스", "ooriintertex"],
    "상일식품": ["상일", "sangil"],
    "미찌푸드": ["미찌", "mizzi"],
    "쌤앤파커스": ["쌤앤파커스", "smpk"],
}


def match_submitter(haystack: str, roster: list[dict]) -> dict | None:
    """파일명·내부경로 문자열에서 제출자를 찾아낸다.

    동명이인(김현우 — 상일식품/쌤앤파커스)이 있으므로
    이름이 여러 명 걸리면 회사 힌트로 좁힌다.
    """
    low = haystack.lower()
    hits = [p for p in roster if p["성명"] in haystack]
    if not hits:
        # 이메일 아이디로도 시도
        for p in roster:
            local = p["이메일"].split("@")[0].lower()
            if local and local in low:
                return p
        return None
    if len(hits) == 1:
        return hits[0]
    # 동명이인 — 회사 힌트로 판별
    for p in hits:
        for hint in COMPANY_HINTS.get(p["회사"], []):
            if hint.lower() in low:
                return p
    return None  # 판별 불가 → 수동 확인 대상


# ─────────────────────────────────────────────────────────────
# 3. 수집
# ─────────────────────────────────────────────────────────────

DOC_KINDS = {
    ".html": "웹서비스·대시보드", ".htm": "웹서비스·대시보드",
    ".xlsx": "스프레드시트", ".xlsm": "스프레드시트", ".xls": "스프레드시트", ".csv": "데이터",
    ".pptx": "발표자료", ".ppt": "발표자료",
    ".docx": "문서", ".doc": "문서", ".hwp": "문서", ".hwpx": "문서", ".pdf": "문서",
    ".py": "스크립트", ".js": "스크립트", ".gs": "스크립트", ".json": "데이터",
    ".png": "이미지", ".jpg": "이미지", ".jpeg": "이미지", ".gif": "이미지",
}


def kind_of(name: str) -> str:
    return DOC_KINDS.get(Path(name).suffix.lower(), "기타")


def human_size(n: int) -> str:
    if n < 1024:
        return f"{n}B"
    if n < 1024 * 1024:
        return f"{n / 1024:.0f}KB"
    return f"{n / 1024 / 1024:.1f}MB"


def collect(input_dir: Path, extract_dir: Path, roster: list[dict]) -> tuple[dict, list[dict]]:
    """산출물을 수집해 제출자별로 묶는다."""
    submissions: dict[str, dict] = {}
    unmatched: list[dict] = []

    def key_of(person: dict) -> str:
        return f"{person['성명']}|{person['회사']}"

    def add(person: dict | None, source: str, files: list[dict], mtime: float):
        entry = {
            "원본": source,
            "파일": files,
            "제출일": datetime.fromtimestamp(mtime).strftime("%Y-%m-%d"),
            "총크기": sum(f["크기"] for f in files),
        }
        if person is None:
            unmatched.append(entry)
            return
        k = key_of(person)
        if k not in submissions:
            submissions[k] = {**person, "제출": []}
        submissions[k]["제출"].append(entry)

    for path in sorted(input_dir.iterdir()):
        if path.name.startswith("_") or path.is_dir():
            continue
        if path.suffix.lower() == ".zip":
            print(f"· 압축 해제: {path.name}")
            try:
                files = extract_zip(path, extract_dir)
            except zipfile.BadZipFile:
                print(f"  ! 손상된 zip — 건너뜀: {path.name}")
                continue
            haystack = path.name + " " + " ".join(f["경로"] for f in files)
            add(match_submitter(haystack, roster), path.name, files, path.stat().st_mtime)
        else:
            files = [{
                "이름": path.name, "경로": path.name,
                "크기": path.stat().st_size, "실제경로": str(path),
            }]
            add(match_submitter(path.name, roster), path.name, files, path.stat().st_mtime)

    return submissions, unmatched


# ─────────────────────────────────────────────────────────────
# 4. HTML 생성 (SECO 문서 표준 + 한글 타이포 + 인쇄 안전)
# ─────────────────────────────────────────────────────────────

CSS = """
:root{
  --seco-pink:#E82076; --seco-gray:#918F8F; --seco-gray-light:#F2F2F2;
  --ink:#1f2330; --sub:#55585f; --mute:#9ca0a8; --line:#e3e4e6;
  --ok:#1A9650; --wait:#C8880A; --bg:#F7F8FA;
}
*{box-sizing:border-box;margin:0;padding:0}
body{
  font-family:'Noto Sans KR','Apple SD Gothic Neo','Malgun Gothic',sans-serif;
  word-break:keep-all; overflow-wrap:break-word; line-height:1.8;
  color:var(--ink); background:var(--bg);
}
p,li,td,th,div,span,a,h1,h2,h3,h4,h5,h6{word-break:keep-all;overflow-wrap:break-word}
.page{max-width:1080px;margin:0 auto;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.conf{background:var(--seco-gray-light);color:#6b6e74;font-size:11px;font-weight:600;
  text-align:center;letter-spacing:.3px;padding:5px 0}
.head{display:flex;justify-content:space-between;align-items:center;
  padding:12px 32px;border-bottom:1px solid var(--line)}
.brand-line{display:flex;align-items:center;gap:11px}
.accent-bar{width:5px;height:22px;background:var(--seco-pink);border-radius:2px}
.affiliate{font-size:19px;font-weight:700;color:var(--ink);letter-spacing:-.3px;white-space:nowrap}
.title-line{display:flex;align-items:baseline;gap:12px;margin-top:7px}
.doc-category{font-size:12px;font-weight:700;color:var(--seco-pink);white-space:nowrap}
.doc-title{font-family:'Noto Serif KR',serif;font-size:17px;font-weight:700;color:var(--ink)}
.head-meta{text-align:right;flex-shrink:0}
.docno{font-size:14px;font-weight:700;color:var(--seco-pink);letter-spacing:.3px}
.meta-sub{font-size:11px;color:var(--mute);margin-top:3px}
.body-wrap{padding:24px 32px 8px}
.section-title{font-family:'Noto Serif KR',serif;font-size:15px;font-weight:700;
  margin:26px 0 12px;padding-left:10px;border-left:4px solid var(--seco-pink)}
.section-title:first-child{margin-top:4px}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.kpi{border:1px solid var(--line);border-radius:10px;padding:14px 16px;background:#fff}
.kpi .lb{font-size:11px;color:var(--sub);font-weight:600}
.kpi .vl{font-size:26px;font-weight:800;margin-top:4px;letter-spacing:-.5px}
.kpi .sb{font-size:10.5px;color:var(--mute);margin-top:2px}
.card{border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin-bottom:12px;background:#fff}
.card-head{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;margin-bottom:4px}
.nm{font-size:15px;font-weight:700}
.co{font-size:11px;font-weight:700;color:var(--seco-pink);background:#FDF0F5;
  padding:2px 8px;border-radius:20px}
.dt{font-size:11px;color:var(--mute);margin-left:auto}
.src{font-size:11px;color:var(--sub);margin-bottom:10px}
.files{display:flex;flex-wrap:wrap;gap:6px}
.f{font-size:11.5px;background:var(--seco-gray-light);border-radius:6px;padding:5px 10px;
  display:inline-flex;align-items:center;gap:6px}
.f .k{font-weight:700;color:var(--seco-pink);font-size:10px}
.f .z{color:var(--mute);font-size:10px}
table{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:4px}
th{background:var(--seco-gray-light);font-weight:700;text-align:left;
  padding:9px 10px;border-bottom:2px solid var(--line);font-size:11.5px}
td{padding:8px 10px;border-bottom:1px solid var(--line)}
.st-ok{color:var(--ok);font-weight:700}
.st-wait{color:var(--wait);font-weight:700}
.note{background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;
  padding:12px 14px;font-size:12px;margin-top:12px}
.foot{border-top:1px solid var(--line);padding:13px 32px;display:flex;
  justify-content:space-between;font-size:11px;color:var(--sub);margin-top:20px}
@page{size:A4;margin:12mm 14mm}
@media print{
  html,body{background:#fff}
  .page{box-shadow:none;margin:0;max-width:none;width:100%}
  .conf{page-break-inside:avoid;break-inside:avoid}
  .head{page-break-inside:avoid;break-inside:avoid;page-break-after:avoid;break-after:avoid}
  .section-title,h1,h2,h3{page-break-after:avoid;break-after:avoid}
  .card,.kpi,.note{page-break-inside:avoid;break-inside:avoid}
  table{page-break-inside:auto}
  thead{display:table-header-group}
  tr{page-break-inside:avoid;break-inside:avoid}
  *{-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important}
}
@media(max-width:720px){.kpi-grid{grid-template-columns:repeat(2,1fr)}}
"""


def esc(s) -> str:
    return html.escape(str(s), quote=True)


def render(roster, submissions, unmatched, docno, session) -> str:
    submitted_keys = set(submissions.keys())
    total = len(roster)
    n_sub = len(submitted_keys)
    n_files = sum(len(e["파일"]) for s in submissions.values() for e in s["제출"])
    total_size = sum(e["총크기"] for s in submissions.values() for e in s["제출"])
    today = datetime.now().strftime("%Y-%m-%d")

    # 제출자 카드
    cards = []
    for s in sorted(submissions.values(), key=lambda x: (x["회사"], x["성명"])):
        for e in s["제출"]:
            files_html = "".join(
                f'<span class="f"><span class="k">{esc(kind_of(f["이름"]))}</span>'
                f'{esc(f["이름"])}<span class="z">{esc(human_size(f["크기"]))}</span></span>'
                for f in e["파일"]
            )
            cards.append(f"""
    <div class="card">
      <div class="card-head">
        <span class="nm">{esc(s['성명'])}</span>
        <span class="co">{esc(s['회사'])}</span>
        <span class="dt">{esc(e['제출일'])} · 파일 {len(e['파일'])}건 · {esc(human_size(e['총크기']))}</span>
      </div>
      <div class="src">원본 · {esc(e['원본'])}</div>
      <div class="files">{files_html}</div>
    </div>""")

    # 전체 명부 표
    rows = []
    for p in roster:
        k = f"{p['성명']}|{p['회사']}"
        if k in submitted_keys:
            s = submissions[k]
            cnt = sum(len(e["파일"]) for e in s["제출"])
            dt = ", ".join(e["제출일"] for e in s["제출"])
            status = f'<span class="st-ok">제출</span>'
            detail = f"{cnt}건"
        else:
            status, detail, dt = '<span class="st-wait">대기</span>', "—", "—"
        rows.append(
            f"<tr><td>{esc(p['성명'])}</td><td>{esc(p['회사'])}</td>"
            f"<td>{esc(p['이메일'])}</td><td>{status}</td><td>{detail}</td><td>{dt}</td></tr>"
        )

    unmatched_html = ""
    if unmatched:
        items = "".join(
            f"<li>{esc(u['원본'])} — 파일 {len(u['파일'])}건 ({esc(human_size(u['총크기']))})</li>"
            for u in unmatched
        )
        unmatched_html = f"""
      <div class="section-title">귀속 미확인 산출물</div>
      <div class="note"><b>{len(unmatched)}건</b>은 파일명에서 제출자를 특정하지 못했다.
      파일명에 성명이 없거나 동명이인(김현우 — 상일식품·쌤앤파커스)이 회사 표기 없이 제출된 경우다.
      아래 항목은 수동으로 귀속 확인이 필요하다.
      <ul style="margin:8px 0 0 18px">{items}</ul></div>"""

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DX과제 산출물 종합 — {esc(session)}</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=Noto+Serif+KR:wght@400;600;700&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<div class="page">
  <div class="conf">내부 한정 · 개인정보(성명) 포함 · 열람·배포 제한</div>
  <div class="head">
    <div class="brand-block">
      <div class="brand-line">
        <span class="accent-bar"></span>
        <span class="affiliate">좋은사람들</span>
      </div>
      <div class="title-line">
        <span class="doc-category">교육 산출물</span>
        <span class="doc-title">AI DX 오프라인 심화캠프 DX과제 산출물 종합</span>
      </div>
    </div>
    <div class="head-meta">
      <div class="docno">{esc(docno)}</div>
      <div class="meta-sub">비제조부문 인사팀 · {esc(today)}</div>
    </div>
  </div>

  <div class="body-wrap">
    <div class="section-title">제출 현황 요약</div>
    <div class="kpi-grid">
      <div class="kpi"><div class="lb">캠프 대상</div><div class="vl">{total}<span style="font-size:14px">명</span></div><div class="sb">8/13 심화캠프 확정 인원</div></div>
      <div class="kpi"><div class="lb">산출물 제출</div><div class="vl" style="color:var(--ok)">{n_sub}<span style="font-size:14px">명</span></div><div class="sb">제출률 {n_sub / total * 100:.0f}%</div></div>
      <div class="kpi"><div class="lb">총 파일</div><div class="vl">{n_files}<span style="font-size:14px">건</span></div><div class="sb">합계 {esc(human_size(total_size))}</div></div>
      <div class="kpi"><div class="lb">미제출</div><div class="vl" style="color:var(--wait)">{total - n_sub}<span style="font-size:14px">명</span></div><div class="sb">독려 필요 인원</div></div>
    </div>

    <div class="section-title">제출자별 산출물</div>
    {"".join(cards) if cards else '<div class="note">수집된 산출물이 없다. --input 폴더에 짚파일을 넣고 다시 실행할 것.</div>'}
{unmatched_html}
    <div class="section-title">대상자 전체 제출 현황 ({total}명)</div>
    <table>
      <thead><tr><th>성명</th><th>회사</th><th>이메일</th><th>상태</th><th>파일</th><th>제출일</th></tr></thead>
      <tbody>{"".join(rows)}</tbody>
    </table>
  </div>

  <div class="foot">
    <div>비제조부문 인사팀</div>
    <div>{esc(docno)} · {esc(today)}</div>
  </div>
</div>
</body>
</html>
"""


# ─────────────────────────────────────────────────────────────
# 5. 진입점
# ─────────────────────────────────────────────────────────────

def main() -> int:
    ap = argparse.ArgumentParser(description="심화캠프 DX과제 산출물 종합")
    ap.add_argument("--input", required=True, help="산출물 원본 폴더")
    ap.add_argument("--extract-dir", default=None, help="압축 해제 위치")
    ap.add_argument("--output", default=None, help="생성할 HTML 경로")
    ap.add_argument("--roster", default=None, help="대상자 명부 JSON")
    ap.add_argument("--docno", default="HR-2026-DX-교육-11", help="문서번호")
    args = ap.parse_args()

    input_dir = Path(args.input).expanduser().resolve()
    if not input_dir.is_dir():
        print(f"오류 — 입력 폴더를 찾을 수 없다: {input_dir}", file=sys.stderr)
        return 1

    roster_path = Path(args.roster) if args.roster else \
        Path(__file__).resolve().parent.parent / "data" / "심화캠프_대상자_23명.json"
    if not roster_path.is_file():
        print(f"오류 — 명부 JSON을 찾을 수 없다: {roster_path}", file=sys.stderr)
        return 1

    meta = json.loads(roster_path.read_text(encoding="utf-8"))
    roster = meta["대상자"]

    extract_dir = Path(args.extract_dir) if args.extract_dir else input_dir / "_해제"
    output = Path(args.output) if args.output else input_dir / "산출물_종합.html"
    extract_dir.mkdir(parents=True, exist_ok=True)

    print(f"입력 : {input_dir}")
    print(f"해제 : {extract_dir}\n")

    submissions, unmatched = collect(input_dir, extract_dir, roster)

    print(f"\n제출자 {len(submissions)}명 / 대상 {len(roster)}명")
    if unmatched:
        print(f"귀속 미확인 {len(unmatched)}건 — HTML 하단에서 확인할 것")

    output.write_text(
        render(roster, submissions, unmatched, args.docno, meta.get("세션", "")),
        encoding="utf-8",
    )
    print(f"\n생성 완료 → {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
