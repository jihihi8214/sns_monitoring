# 정관계/AI-ICT 인사 SNS 모니터링 (테스트: 대통령)

기존 `HANDOFF.md` (정부 보도자료/인사소식 모니터링) 구조를 참고해서 만든 X/페이스북 버전입니다.
API 없이 브라우저(Playwright)로 공개 페이지를 직접 읽는 방식이라, 발급 비용은 없지만 사이트 구조 변경/로그인 요구에 취약할 수 있습니다.

## 1. 로컬 준비 (최초 1회)

```bash
cd sns_monitor
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## 2. 이메일 설정

`.env.example`을 복사해 `.env`를 만들고 값을 채우세요.

```bash
cp .env.example .env
```

- `EMAIL_TO`: 받는 사람 (jane.kkim@kakaocorp.com)
- `SMTP_*`: 보내는 계정 정보. Gmail이면 앱 비밀번호 필요 (구글 계정 → 보안 → 앱 비밀번호).

`.env`는 절대 외부에 공유하지 마세요. 전달할 땐 `.env.example`만 공유합니다.

## 3. (필요 시) X/페이스북 로그인 세션 저장

X, 페이스북 모두 비로그인 상태로는 최신 글이 안 보이거나 일부만 보일 수 있습니다.
아래 스크립트로 한 번만 직접 로그인해두면 이후 자동 실행 시 로그인 상태를 재사용합니다.

```bash
python3 scripts/login_setup.py twitter
python3 scripts/login_setup.py facebook
```

브라우저 창이 뜨면 본인이 직접 아이디/비밀번호를 입력해서 로그인하고, 터미널로 돌아와 Enter를 누르면 세션이 `data/`에 저장됩니다.

## 4. 대상 계정 (현재는 테스트로 대통령만 등록)

`sources.json`에서 관리합니다.

```json
{
  "name": "대통령(이재명)",
  "twitter": ["Jaemyung_Lee", "KOREA"],
  "facebook": ["jaemyunglee"]
}
```

이후 여/야 당대표, 원내대표, 과방위원, 배경훈, 송경희, 김종철 등을 추가하려면 이 파일에 인물 객체를 추가하면 됩니다.

## 5. 실행

```bash
python3 monitor.py
```

- 새 게시물이 있으면 `jane.kkim@kakaocorp.com`으로 메일 발송 + `reports/`에 기록 저장
- 새 게시물이 없으면 메일 발송 안 함
- 이미 확인한 게시물은 `data/seen.json`에 기록되어 중복 알림 방지

## 6. Codex 앱 예약 설정 (매일 아침 자동 실행)

기존 정부 모니터링 에이전트와 동일하게 Codex 앱의 "예약됨" 기능을 사용하세요.

- 예약 이름: `정관계/AI-ICT SNS 모니터링`
- 실행 주기: 매일 1회 (예: 오전 8시) — 브라우저 스크래핑 방식이라 시간당 실행처럼 자주 돌리면 차단 위험이 있어 하루 1~2회를 권장합니다.
- 실행 환경: local
- 작업 폴더: 이 `sns_monitor` 폴더 경로
- 실행 명령: `venv/bin/python3 monitor.py`

## 7. 알려진 제약

- X: 비로그인 상태에서는 트윗이 일부만 보이거나 로그인 요구가 뜰 수 있습니다. `login_setup.py`로 세션을 저장해두는 걸 권장합니다.
- 페이스북: 개인 프로필은 원천적으로 접근이 막혀 있어 대상이 될 수 없습니다. 공개 "페이지"만 대상 가능합니다.
- 두 플랫폼 모두 HTML 구조가 바뀌면 `monitor.py`의 선택자(selector)를 손봐야 할 수 있습니다.
- 현재 요약은 게시물 앞부분 텍스트를 그대로 보여주는 수준입니다. AI 요약이 필요하면 Codex 실행 단계에서 새 게시물 텍스트를 요약하도록 프롬프트를 추가하는 방식을 권장합니다.

## 8. 다음 확장 계획

- `sources.json`에 여/야 당대표, 원내대표, 과방위원, 배경훈, 송경희, 김종철 등 추가
- 대통령 테스트가 안정적으로 동작 확인되면 순차 확대
