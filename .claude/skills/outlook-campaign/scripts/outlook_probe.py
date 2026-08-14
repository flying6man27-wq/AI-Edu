# -*- coding: utf-8 -*-
"""아웃룩 환경 진단 — 이 스킬의 첫 단계. 읽기만 하고 아무것도 바꾸지 않는다.

이걸 먼저 돌리는 이유: 데스크톱 Outlook 이 있는지 없는지에 따라 발송·수집 경로가
완전히 갈리는데, 사용자가 그걸 모르는 경우가 대부분이다. 추측하지 말고 확인한다.

    python outlook_probe.py
"""
import sys

sys.path.insert(0, __file__.rsplit("\\", 1)[0].rsplit("/", 1)[0])
from _outlook_com import OL_FOLDER_INBOX, OL_FOLDER_SENT, OutlookUnavailable, connect  # noqa: E402


def main():
    print("=" * 60)
    print("아웃룩 환경 진단")
    print("=" * 60)

    if not sys.platform.startswith("win"):
        print(f"[X] 현재 OS: {sys.platform} — 데스크톱 Outlook COM 은 Windows 전용이다.")
        print("    → references/web-fallback.md 의 웹 경로로 진행할 것.")
        return 2

    try:
        app, ns = connect()
    except OutlookUnavailable as e:
        print(f"[X] {e}")
        print("    → references/web-fallback.md 의 웹 경로로 진행할 것.")
        return 2

    print("[O] 데스크톱 Outlook COM 연결 성공")
    try:
        print(f"    버전: {app.Version}")
    except Exception:
        pass

    # 어떤 계정이 붙어 있는지 — 여러 계정이면 어느 주소로 나가는지가 중요해진다
    try:
        accounts = [ns.Accounts.Item(i).SmtpAddress for i in range(1, int(ns.Accounts.Count) + 1)]
        print(f"    계정 {len(accounts)}개: {', '.join(a for a in accounts if a)}")
        if len(accounts) > 1:
            print("    [!] 계정이 여러 개다. 발송 시 어느 주소로 나갈지 사용자에게 확인할 것.")
    except Exception:
        print("    [!] 계정 목록을 읽지 못했다.")

    for fid, label in ((OL_FOLDER_INBOX, "받은편지함"), (OL_FOLDER_SENT, "보낸편지함")):
        try:
            f = ns.GetDefaultFolder(fid)
            print(f"    {label}: {f.Items.Count}건 접근 가능")
        except Exception as e:
            print(f"    [!] {label} 접근 실패: {e}")

    print()
    print("판정: COM 경로 사용 가능 — outlook_send.py / outlook_collect.py 를 그대로 쓸 것.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
