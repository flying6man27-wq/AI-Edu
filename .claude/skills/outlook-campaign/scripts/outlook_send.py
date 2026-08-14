# -*- coding: utf-8 -*-
"""캠페인 발송 — 현황 CSV + 본문 템플릿 → Outlook 초안 생성(기본) 또는 발송.

기본값이 '초안'인 게 이 스크립트의 핵심 설계다. 메일은 취소가 안 되고, 명단
20명짜리 오발송은 하루를 통째로 태운다. 사람이 초안 한 건을 눈으로 보고 나서
--send 를 붙이는 흐름이 비용 대비 가장 안전하다.

    # 1) 초안 1건만 만들어 눈으로 확인
    python outlook_send.py --csv 현황.csv --template 본문.html --subject "[SECO] DX심화캠프 과제 안내" --limit 1

    # 2) 확인됐으면 전체 발송
    python outlook_send.py --csv 현황.csv --template 본문.html --subject "..." --send

본문 템플릿의 {컬럼명} 은 CSV 같은 행의 값으로 치환된다. 예: {성명}, {회사}, {과제명}
"""
import argparse
import csv
import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _outlook_com import OL_MAIL_ITEM, OutlookUnavailable, connect  # noqa: E402

COL_SENT = "이메일발송일"
COL_STATUS = "상태"
COL_EMAIL = "이메일"


def read_rows(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return list(r), list(r.fieldnames or [])


def write_rows(path, rows, fields):
    for extra in (COL_SENT, COL_STATUS):
        if extra not in fields:
            fields.append(extra)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def render(template, row):
    """{컬럼명} 을 행 값으로 치환. 남은 미치환 자리표시자는 눈에 띄게 보고한다.

    조용히 빈칸으로 두면 '{성명}님 안녕하세요' 가 그대로 나가는 사고가 난다.

    단, 미치환 검사에서 HTML 주석과 <style> 블록은 제외한다. 템플릿 주석에 적어둔
    사용법('{컬럼명} 은 치환된다')이나 CSS 중괄호를 누락으로 잡으면 정상 템플릿이
    영영 발송되지 않는다 — 실제로 이 스킬의 예시 템플릿이 그렇게 걸렸다.
    """
    out = template
    for k, v in row.items():
        if k:
            out = out.replace("{" + k + "}", str(v or ""))

    scan = re.sub(r"<!--.*?-->", "", out, flags=re.S)
    scan = re.sub(r"<style\b.*?</style>", "", scan, flags=re.S | re.I)
    leftover = sorted(set(re.findall(r"\{([가-힣A-Za-z_][가-힣A-Za-z0-9_]*)\}", scan)))
    return out, leftover


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", required=True, help="현황 CSV (명단 + 진행상태)")
    ap.add_argument("--template", required=True, help="본문 파일 (.html 또는 .txt)")
    ap.add_argument("--subject", required=True, help="제목. {컬럼명} 치환 가능")
    ap.add_argument("--attach", action="append", default=[], help="첨부할 파일 경로(반복 가능)")
    ap.add_argument("--send", action="store_true", help="실제 발송. 없으면 초안만 생성")
    ap.add_argument("--limit", type=int, default=0, help="상위 N명만 처리(파일럿용)")
    ap.add_argument("--resend", action="store_true", help="이미 발송된 행도 다시 처리")
    ap.add_argument("--only", default="", help="특정 이메일만 처리(쉼표구분)")
    args = ap.parse_args()

    rows, fields = read_rows(args.csv)
    with open(args.template, encoding="utf-8") as f:
        template = f.read()
    is_html = args.template.lower().endswith((".html", ".htm"))

    for p in args.attach:
        if not os.path.exists(p):
            print(f"[중단] 첨부 파일이 없다: {p}")
            return 1

    only = {e.strip().lower() for e in args.only.split(",") if e.strip()}

    try:
        app, _ = connect()
    except OutlookUnavailable as e:
        print(f"[중단] {e}")
        print("→ references/web-fallback.md 참조.")
        return 2

    today = datetime.date.today().isoformat()
    done = skipped = 0

    for row in rows:
        email = (row.get(COL_EMAIL) or "").strip()
        name = row.get("성명") or row.get("이름") or email

        if not email or "@" not in email:
            print(f"  [건너뜀] {name}: 이메일 없음/형식오류 → 사람이 확인해야 한다")
            skipped += 1
            continue
        if only and email.lower() not in only:
            continue
        if row.get(COL_SENT) and not args.resend:
            skipped += 1
            continue
        if args.limit and done >= args.limit:
            break

        body, leftover = render(template, row)
        subject, sub_left = render(args.subject, row)
        if leftover or sub_left:
            # 치환 안 된 자리표시자를 안고 보내면 그대로 수신자 화면에 뜬다
            print(f"  [중단] {name}: 치환되지 않은 항목 {sorted(set(leftover + sub_left))}")
            print("         CSV 컬럼명과 템플릿 자리표시자가 일치하는지 확인할 것.")
            return 1

        mail = app.CreateItem(OL_MAIL_ITEM)
        mail.To = email
        mail.Subject = subject
        if is_html:
            mail.HTMLBody = body
        else:
            mail.Body = body
        for p in args.attach:
            mail.Attachments.Add(os.path.abspath(p))

        if args.send:
            mail.Send()
            row[COL_SENT] = today
            row[COL_STATUS] = "발송완료"
            print(f"  [발송] {name} <{email}>")
        else:
            mail.Save()  # 임시보관함
            print(f"  [초안] {name} <{email}>")
        done += 1

    if args.send:
        write_rows(args.csv, rows, fields)

    mode = "발송" if args.send else "초안 생성"
    print(f"\n{mode} {done}건 · 건너뜀 {skipped}건")
    if not args.send:
        print("초안은 Outlook 임시보관함에 있다. 내용을 확인한 뒤 --send 로 재실행할 것.")
        print("(초안은 CSV 를 갱신하지 않는다 — 실제로 나간 것만 기록으로 남긴다)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
