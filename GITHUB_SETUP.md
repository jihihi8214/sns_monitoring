# GitHub Actions로 완전 자동화하기

노트북/Chrome 상태와 무관하게 매시 정각 자동으로 돌아가도록 GitHub 서버에서 실행하는 방식입니다. 무료입니다 (GitHub Actions 무료 사용량 안에서 충분히 커버됨).

## 0. 이 폴더 구조

리포지토리 루트 = 지금 이 `sns_monitor` 폴더의 "내용물"이어야 합니다 (폴더 자체를 통째로 올리는 게 아니라, 안에 있는 파일들을 리포지토리 최상단에 둡니다). 즉 GitHub에 올렸을 때 이렇게 보여야 합니다.

```
(리포지토리 루트)
├── .github/workflows/monitor.yml
├── monitor.py
├── sources.json
├── requirements.txt
├── data/
├── reports/
└── scripts/
```

## 1. 새 GitHub 리포지토리 만들기

1. github.com 로그인 → 우측 상단 "+" → "New repository"
2. 이름 예: `sns-monitor` (아무 이름이나 가능)
3. Private로 설정 (권장 — 공개 안 해도 됨)
4. "Create repository" 클릭

## 2. 로컬에서 파일 업로드

터미널에서:

```bash
cd "/Users/jane.kkim/Library/Application Support/Claude/local-agent-mode-sessions/f2ebfed1-3ee2-4c17-b6c8-c7b4dfdfbff8/481257bc-a931-406e-8b06-11424a6a53dd/local_808ade1b-a408-4854-9f15-dad9541e3f09/outputs/sns_monitor"

git init
git add .
git commit -m "초기 설정"
git branch -M main
git remote add origin https://github.com/본인아이디/sns-monitor.git
git push -u origin main
```

(`.env` 파일은 이미 없는 상태일 거예요 — 있다면 절대 올리지 말고 삭제 후 진행하세요. 안전을 위해 `.gitignore`에 `.env`를 추가해두는 걸 권장해요.)

## 3. 이메일 발송 계정 준비 (앱 비밀번호)

Gmail 기준: 구글 계정 → 보안 → 2단계 인증 켜기 → "앱 비밀번호"에서 16자리 코드 발급.

## 4. X/페이스북 로그인 세션 만들기 (로컬에서 1회만)

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium

python3 scripts/login_setup.py twitter
python3 scripts/login_setup.py facebook
```

브라우저가 뜨면 직접 로그인 → 터미널에서 Enter. 그러면 `data/twitter_state.json`, `data/facebook_state.json`이 생깁니다.

이 파일들을 base64로 인코딩합니다 (다음 단계 GitHub Secrets에 넣기 위함):

```bash
base64 -i data/twitter_state.json | pbcopy
```

(맥에서 `pbcopy`를 쓰면 결과가 바로 클립보드에 복사됩니다. 이걸 붙여넣을 곳은 다음 단계입니다.)

## 5. GitHub Secrets 등록

리포지토리 페이지 → Settings → Secrets and variables → Actions → "New repository secret"

아래 항목을 하나씩 등록하세요 (이름은 정확히 일치해야 합니다).

| Secret 이름 | 값 |
|---|---|
| `EMAIL_TO` | jane.kkim@kakaocorp.com |
| `SMTP_HOST` | smtp.gmail.com |
| `SMTP_PORT` | 587 |
| `SMTP_USERNAME` | 보내는 계정 이메일 |
| `SMTP_PASSWORD` | 4단계에서 발급받은 앱 비밀번호 |
| `SMTP_FROM` | 보내는 계정 이메일 (USERNAME과 동일) |
| `TWITTER_STATE_B64` | 4단계에서 base64 인코딩한 twitter_state.json 내용 |
| `FACEBOOK_STATE_B64` | 4단계에서 base64 인코딩한 facebook_state.json 내용 |

이 값들은 GitHub 자체 암호화 저장소에 들어가고, 저(Claude)를 포함해 아무도 다시 볼 수 없습니다.

## 6. 확인

Settings → Actions → "General"에서 Actions가 활성화되어 있는지 확인 후, 리포지토리의 "Actions" 탭 → "SNS Monitor" 워크플로 → "Run workflow" 버튼으로 수동 실행해서 정상 작동하는지 먼저 테스트해보세요.

정상 작동하면 이후로는 매시 정각 자동 실행됩니다 (GitHub 부하에 따라 몇 분 지연될 수 있음).

## 7. 알아둘 점

- 이 방식은 Claude(저)가 실행 시점에 관여하지 않는 순수 스크립트라서, 게시글 "요약"은 AI가 다시 쓴 요약이 아니라 원문 앞부분 그대로(최대 300자)가 들어갑니다.
- 서버 IP에서 접속하는 방식이라 X/페이스북이 차단하거나 로그인 세션이 만료될 수 있습니다. 이 경우 4단계를 다시 실행해서 Secrets를 갱신하세요.
- 대상 계정 추가는 `sources.json`을 수정 후 다시 `git push`하면 반영됩니다.
