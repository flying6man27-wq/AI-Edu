# -*- coding: utf-8 -*-
"""Outlook 데스크톱(COM) 공통 헬퍼 — 발송·수집 스크립트가 함께 쓴다.

여기 모아둔 것들은 전부 "그냥 하면 틀리는" 지점들이다:
  · Exchange 계정의 SenderEmailAddress 는 SMTP 주소가 아니라 X500 DN 이다.
  · 회신 메일의 첨부에는 상대방 서명 로고가 섞여 들어온다.
  · Restrict() 날짜 필터는 Windows 로캘을 탄다.
각 함수 주석에 왜 그렇게 처리하는지 적어뒀다.
"""
import datetime
import os
import re

# MAPI 프로퍼티 태그 — PropertyAccessor 로만 읽을 수 있는 값들
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001E"

OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_MAIL_ITEM = 0


class OutlookUnavailable(RuntimeError):
    """데스크톱 Outlook COM 에 붙을 수 없을 때. 호출부는 이걸 잡아 웹 경로를 안내한다."""


def connect():
    """Outlook.Application 에 연결하고 (app, namespace) 를 돌려준다.

    Dispatch 는 Outlook 이 꺼져 있으면 자동으로 띄운다. 이미 떠 있으면 그 인스턴스를
    재사용하므로 사용자가 보고 있는 화면과 같은 프로필을 쓴다 — 이게 중요하다.
    다른 프로필을 열면 보낸 메일이 사용자 눈에 안 보인다.
    """
    try:
        import win32com.client  # noqa: F401  (pywin32)
    except ImportError as e:
        raise OutlookUnavailable(
            "pywin32 가 없다. Windows + 데스크톱 Outlook 환경에서 `pip install pywin32` 후 재시도."
        ) from e

    import win32com.client

    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        ns.Logon("", "", False, False)
    except Exception as e:
        raise OutlookUnavailable(f"Outlook COM 연결 실패: {e}") from e
    return app, ns


def smtp_of(item):
    """메일 항목의 발신자 SMTP 주소를 구한다.

    사내 Exchange 발신자는 SenderEmailAddress 가
    '/O=EXCHANGELABS/OU=.../CN=RECIPIENTS/CN=abc123' 같은 X500 DN 으로 온다.
    이걸 그대로 명단의 이메일과 비교하면 회신이 단 한 건도 매칭되지 않는다.
    PropertyAccessor 로 PR_SMTP_ADDRESS 를 읽어야 진짜 주소가 나온다.
    """
    raw = ""
    try:
        raw = (item.SenderEmailAddress or "").strip()
    except Exception:
        pass

    if raw and not raw.upper().startswith("/O="):
        return raw.lower()

    # X500 DN 이거나 비어 있으면 SMTP 프로퍼티를 직접 읽는다
    for getter in (
        lambda: item.Sender.GetExchangeUser().PrimarySmtpAddress,
        lambda: item.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS),
        lambda: item.Sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS),
    ):
        try:
            v = (getter() or "").strip()
            if v and "@" in v:
                return v.lower()
        except Exception:
            continue
    return raw.lower()


def is_signature_image(att):
    """상대 서명에 박힌 로고·아이콘인가?

    회신 메일을 첨부 기준으로만 긁으면 회사 로고 gif, 트위터 아이콘, 배경 이미지가
    과제 파일과 같은 폴더에 우수수 쌓인다. 20명한테 받으면 순식간에 수백 개다.
    걸러내는 기준은 두 가지 — 본문에 박힌 인라인 이미지는 Content-ID 를 갖는다는 것,
    그리고 서명 그래픽은 거의 예외 없이 작다는 것.
    """
    name = (getattr(att, "FileName", "") or "").lower()

    # 1) 인라인(CID) 이미지 — 본문에 박힌 것이지 첨부한 게 아니다
    try:
        cid = att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
        if cid:
            return True
    except Exception:
        pass

    # 2) image* / oledata 류 자동생성 이름
    if re.match(r"^(image\d+|oledata|ole\d+|~wr)", name):
        return True

    # 3) 작은 그림 파일 — 실제 제출물이 40KB 미만인 경우는 사실상 없다
    if name.endswith((".gif", ".png", ".jpg", ".jpeg", ".bmp", ".emf", ".wmf", ".svg")):
        try:
            if int(att.Size) < 40 * 1024:
                return True
        except Exception:
            return True
    return False


def safe_filename(s, maxlen=120):
    """Windows 파일명으로 못 쓰는 문자를 정리한다. 한글은 그대로 둔다."""
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(s or "")).strip(" .")
    return (s[:maxlen] or "untitled")


def save_attachments(item, outdir, prefix="", skip_signature=True):
    """메일의 첨부를 outdir 에 저장하고 저장된 파일명 리스트를 돌려준다.

    같은 이름이 이미 있으면 덮어쓰지 않고 _2, _3 을 붙인다 — 두 사람이 똑같이
    '과제.docx' 로 보내오는 일이 흔하고, 덮어쓰면 앞사람 제출물이 조용히 사라진다.
    """
    os.makedirs(outdir, exist_ok=True)
    saved = []
    try:
        n = int(item.Attachments.Count)
    except Exception:
        return saved

    for i in range(1, n + 1):
        att = item.Attachments.Item(i)
        if skip_signature and is_signature_image(att):
            continue
        base = safe_filename(getattr(att, "FileName", "") or f"attachment_{i}")
        stem, ext = os.path.splitext(base)
        fname = f"{prefix}{stem}{ext}" if prefix else base
        path = os.path.join(outdir, fname)
        k = 2
        while os.path.exists(path):
            path = os.path.join(outdir, f"{prefix}{stem}_{k}{ext}")
            k += 1
        try:
            att.SaveAsFile(path)
            saved.append(os.path.basename(path))
        except Exception as e:
            print(f"    [경고] 첨부 저장 실패 {base}: {e}")
    return saved


def recent_items(folder, since_days=30):
    """폴더에서 최근 N일 항목을 최신순으로 준다.

    Restrict() 의 날짜 문자열은 Windows 지역설정을 타서 한국어 환경에서 조용히
    0건을 반환하는 일이 있다. 그래서 Restrict 를 시도하되, 실패하거나 0건이면
    정렬 후 직접 훑는 방식으로 넘어간다 — 느리지만 절대 놓치지 않는다.
    """
    cutoff = datetime.datetime.now() - datetime.timedelta(days=since_days)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # 최신순

    try:
        flt = "[ReceivedTime] >= '" + cutoff.strftime("%m/%d/%Y %H:%M %p") + "'"
        r = items.Restrict(flt)
        if int(r.Count) > 0:
            return list(_iter_com(r))
    except Exception:
        pass

    out = []
    for it in _iter_com(items):
        try:
            rt = it.ReceivedTime
            # pywin32 datetime 은 tz-aware 로 올 수 있어 naive 로 맞춘다
            rt = datetime.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
            if rt < cutoff:
                break  # 최신순 정렬이므로 여기서부터는 전부 오래된 것
        except Exception:
            continue
        out.append(it)
    return out


def _iter_com(collection):
    """COM 컬렉션을 예외에 강하게 순회한다. 손상된 항목 하나가 전체를 멈추지 않도록."""
    try:
        n = int(collection.Count)
    except Exception:
        return
    for i in range(1, n + 1):
        try:
            yield collection.Item(i)
        except Exception:
            continue
