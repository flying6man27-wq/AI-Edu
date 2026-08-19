# 데스크톱 Outlook 이 없을 때 — 웹 경로

`outlook_probe.py` 가 COM 연결에 실패했을 때만 여기로 온다. 웹 경로는 **발송은 되지만
수집은 반자동**이다. 이 비대칭을 사용자에게 먼저 알리고 시작할 것 — "다 자동으로 되겠지"
하고 20명 캠페인을 시작했다가 회수 단계에서 손으로 다 받아야 한다는 걸 알게 되면
그때는 이미 늦다.

## 발송 — 본문 div 직접 주입

Outlook 웹(`outlook.cloud.microsoft/mail`)에서 서식 있는 본문을 넣을 때
`navigator.clipboard.write()` 를 브라우저 자동화로 호출하면 `NotAllowedError:
Document is not focused` 로 거의 항상 실패한다. 자동화 컨텍스트를 브라우저가
"포커스됨"으로 보지 않기 때문이다. 클립보드를 우회해 본문에 직접 넣는다.

1. 새 메일 → 받는사람 입력 → Enter 로 확정 → 제목 입력.
2. **본문 영역을 실제로 한 번 클릭한다.** `document.hasFocus()` 를 true 로 만드는
   단계이고, 이걸 건너뛰면 아래 스크립트도 조용히 실패한다.
3. HTML 은 base64 로 넘겨 따옴표·개행 문제를 피한다:

```js
const body = document.querySelector('[aria-label="메시지 본문"]');
body.focus();
const sel = window.getSelection();
const rg = document.createRange();
rg.selectNodeContents(body);
sel.removeAllRanges(); sel.addRange(rg);
document.execCommand('insertHTML', false, html);
body.dispatchEvent(new InputEvent('input', {bubbles:true, inputType:'insertFromPaste'}));
```

4. 발송 전 DOM 을 직접 읽어 검증한다(스크린샷은 지연 렌더링 때문에 타임아웃이 잦다):

```js
({
  to: document.querySelector('[aria-label="받는 사람"]').innerText,
  subject: document.querySelector('[aria-label="과목"]').value,
  links: Array.from(document.querySelector('[aria-label="메시지 본문"]')
           .querySelectorAll('a')).map(a => a.getAttribute('href')),
})
```

5. 보내기는 `aria-label="보내기"` 버튼.

## 수집 — 자동화하지 말 것

웹에서 첨부를 내려받게 하는 스크립트는 만들지 않는다. 브라우저 다운로드는 샌드박스
정책·파일 대화상자에 막히고, 어렵게 우회해도 어느 파일이 누구 것인지 잃어버린다.
잘못 붙은 이름으로 저장된 제출물은 나중에 아무도 못 고친다.

대신 사람이 하도록 돕는다:

1. 웹에서 캠페인 제목으로 검색한다 — `subject:"DX심화캠프" hasattachment:yes`
2. 사용자가 첨부를 한 폴더에 내려받는다.
3. 그 폴더를 스킬이 읽어 현황 CSV 를 갱신한다:

```
python scripts/outlook_collect.py --csv 현황.csv --outdir 내려받은폴더 --from-folder
```

> `--from-folder` 는 아직 구현돼 있지 않다. 필요해지는 시점에 추가할 것 —
> 파일명에서 이름을 역추적하는 규칙이 조직마다 달라, 실제 회신 파일 몇 건을
> 보고 나서 만드는 편이 정확하다. 그전까지는 사용자가 파일을 보며 CSV 의
> `회신일`·`수집파일`을 채우는 걸 도와주면 된다.

## 근본 해결책 — Microsoft 365 커넥터

이 조직 계정에는 Microsoft 365 커넥터가 이미 설치돼 있으나 대화에서 꺼져 있다
(`ListConnectors` 로 확인 가능: `enabledInChat: false`). 켜면 Graph API 로 발송·회신
조회·첨부 다운로드가 전부 되므로 COM 없이도 완전 자동화가 된다.

- claude.ai → 설정 → 커넥터 → Microsoft 365 활성화
- 회사 테넌트라면 "관리자 승인 필요" 가 뜰 수 있다. 그 경우 IT 요청이 필요하고,
  승인까지는 COM 경로(로컬 PC)가 가장 빠른 길이다.
