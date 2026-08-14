# -*- coding: utf-8 -*-
"""회신 수집 — 받은편지함을 훑어 회신을 찾고, 첨부를 저장하고, 현황 CSV를 갱신한다.

    python outlook_collect.py --csv 현황.csv --outdir 수집물 --subject-contains "DX심화캠프" --dry-run
    python outlook_collect.py --csv 현황.csv --outdir 수집물 --subject-contains "DX심화캠프"

--dry-run 을 먼저 돌리는 습관을 들일 것. 무엇이 누구 것으로 잡혔는지 확인하고 나서
실제로 저장한다. 잘못 매칭된 파일이 남의 이름으로 저장되면 나중에 되돌리기 어렵다.
"""
import argparse
import csv
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _outlook_com import (  # noqa: E402
    OL_FOLDER_INBOX,
    OutlookUnavailable,
    connect,
    recent_items,
    safe_filename,
    save_attachments,
    smtp_of,
)

COL_EMAIL, COL_REPLY, COL_STATUS, COL_FILES, COL_NOTE = (
    "이메일", "회신일", "상태", "수집파일", "비고",
)


def read_rows(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return list(r), list(r.fieldnames or [])


def write_rows(path, rows, fields):
    for extra in (COL_REPLY, COL_STATUS, COL_FILES, COL_NOTE):
        if extra not in fields:
            fields.append(extra)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def is_bounce(item):
    """반송(NDR)인가 — 주소가 틀렸다는 뜻이므로 회신과 전혀 다르게 취급해야 한다.

    추정 주소로 보낸 캠페인에서는 반송이 반드시 나온다. 이걸 '회신 없음'으로
    뭉뚱그리면 그 사람은 영영 독촉만 받고 메일은 계속 허공으로 나간다.
    """
    try:
        cls = (item.MessageClass or "")
        if cls.startswith("REPORT.") and "NDR" in cls.upper():
            return True
    except Exception:
        pass
    try:
        s = (smtp_of(item) or "")
        return any(k in s for k in ("postmaster", "mailer-daemon", "microsoftexchange"))
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", required=True)
    ap.add_argument("--outdir", required=True, help="첨부 저장 폴더")
    ap.add_argument("--subject-contains", default="", help="이 문자열이 제목에 있는 메일만 회신으로 본다")
    ap.add_argument("--since-days", type=int, default=30)
    ap.add_argument("--dry-run", action="store_true", help="저장·기록 없이 무엇이 잡히는지만 출력")
    args = ap.parse_args()

    rows, fields = read_rows(args.csv)
    by_email = {}
    for r in rows:
        e = (r.get(COL_EMAIL) or "").strip().lower()
        if e:
            by_email[e] = r

    try:
        _, ns = connect()
    except OutlookUnavailable as e:
        print(f"[중단] {e}")
        print("→ references/web-fallback.md 참조.")
        return 2

    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = recent_items(inbox, args.since_days)
    print(f"받은편지함 최근 {args.since_days}일 · {len(items)}건 스캔\n")

    kw = args.subject_contains.strip()
    today = datetime.date.today().isoformat()
    matched = bounced = 0
    unmatched = []

    for it in items:
        try:
            subject = it.Subject or ""
        except Exception:
            continue
        if kw and kw not in subject:
            continue

        sender = smtp_of(it)

        if is_bounce(it):
            # 반송 본문에는 원래 수신자 주소가 들어 있다 — 명단과 대조해 표시
            body = ""
            try:
                body = (it.Body or "")[:4000]
            except Exception:
                pass
            for e, row in by_email.items():
                if e in body.lower():
                    row[COL_STATUS] = "반송"
                    row[COL_NOTE] = ((row.get(COL_NOTE) or "") + f" 반송({today}) 주소확인필요").strip()
                    print(f"  [반송] {row.get('성명','?')} <{e}> — 주소가 틀렸다")
                    bounced += 1
                    break
            continue

        row = by_email.get(sender)
        if row is None:
            unmatched.append((sender, subject))
            continue

        name = row.get("성명") or row.get("이름") or sender
        company = row.get("회사") or row.get("계열사") or ""
        prefix = safe_filename(f"{name}_{company}_") if company else safe_filename(f"{name}_")

        if args.dry_run:
            try:
                names = [
                    it.Attachments.Item(i).FileName
                    for i in range(1, int(it.Attachments.Count) + 1)
                ]
            except Exception:
                names = []
            print(f"  [매칭] {name} <{sender}> · 첨부후보 {names}")
            matched += 1
            continue

        saved = save_attachments(it, args.outdir, prefix=prefix)
        row[COL_REPLY] = today
        if saved:
            row[COL_STATUS] = "회신완료"
            prev = [x for x in (row.get(COL_FILES) or "").split("; ") if x]
            row[COL_FILES] = "; ".join(prev + saved)
            print(f"  [수집] {name} · {len(saved)}개 — {', '.join(saved)}")
        else:
            # 회신은 왔는데 첨부가 없다. 본문에 링크로 냈거나 "곧 보내겠습니다" 류다.
            row[COL_STATUS] = "회신(첨부없음)"
            print(f"  [회신] {name} · 첨부 없음 — 본문 확인 필요")
        matched += 1

    if not args.dry_run:
        write_rows(args.csv, rows, fields)

    print(f"\n매칭 {matched}건 · 반송 {bounced}건 · 미매칭 {len(unmatched)}건")

    if unmatched:
        # 미매칭은 버리면 안 된다. 팀장이 대신 회신하거나 개인메일로 보내는 일이 흔하다.
        print("\n[미매칭] 제목은 맞는데 명단에 없는 주소에서 온 회신 — 사람이 판단할 것:")
        for s, subj in unmatched[:20]:
            print(f"    {s} · {subj[:60]}")
        print("    → 본인 대신 보낸 것이면 해당 행 이메일을 고치고 다시 실행하면 잡힌다.")

    still = [r for r in rows if not (r.get(COL_REPLY) or "").strip()
             and (r.get(COL_STATUS) or "") != "반송"]
    if still:
        print(f"\n[미회신] {len(still)}명: " + ", ".join(
            (r.get("성명") or r.get(COL_EMAIL) or "?") for r in still[:30]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
