"""
정관계/AI-ICT 인사 SNS(X, 페이스북) 모니터링 스크립트

- sources.json에 등록된 인물의 X/페이스북 공개 게시물을 확인합니다.
- 이전 실행 대비 새 게시물이 있으면 이메일로 알려줍니다.
- API를 쓰지 않고 브라우저(Playwright)로 공개 페이지를 직접 읽는 방식이라
  플랫폼의 HTML 구조가 바뀌면 선택자(selector)를 수정해야 할 수 있습니다.
- 로그인이 필요한 경우 scripts/login_setup.py를 먼저 실행해 세션을 저장해두세요.

실행:
    python3 monitor.py
"""

import os
import re
import csv
import json
import time
import hashlib
import html
import smtplib
import ssl
import urllib.request
import urllib.error
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from openpyxl import Workbook
from playwright.sync_api import sync_playwright

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
REPORTS_DIR = os.path.join(BASE, "reports")
SEEN_PATH = os.path.join(DATA_DIR, "seen.json")
SOURCES_PATH = os.path.join(BASE, "sources.json")
ENV_PATH = os.path.join(BASE, ".env")
CSV_PATH = os.path.join(DATA_DIR, "new_items.csv")
CSV_HEADER = "확인시각,인물,플랫폼,게시시각표기,내용요약,링크\n"
EXCEL_PATH = os.path.join(DATA_DIR, "sns_monitoring.xlsx")
# GitHub Pages가 이 폴더(docs/)를 그대로 고정 URL로 서빙한다.
# 매 실행마다 이 파일을 덮어써서 커밋/푸시하면, 링크는 그대로인 채 내용만 갱신된다.
DOCS_DIR = os.path.join(BASE, "docs")
HTML_PATH = os.path.join(DOCS_DIR, "index.html")

# 봇 탐지를 조금이라도 줄이기 위한 일반 브라우저 흉내용 컨텍스트 옵션
CONTEXT_ARGS = {
    "user_agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "viewport": {"width": 1366, "height": 900},
    "locale": "ko-KR",
}


# ---------- 설정 로드 ----------

def load_env(path):
    env = {}
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"{path} 이 없습니다. .env.example을 복사해서 .env를 만들고 값을 채워주세요."
        )
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env


def load_sources():
    with open(SOURCES_PATH, "r", encoding="utf-8") as f:
        return json.load(f)["persons"]


def load_seen():
    if os.path.exists(SEEN_PATH):
        with open(SEEN_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_seen(seen):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(SEEN_PATH, "w", encoding="utf-8") as f:
        json.dump(seen, f, ensure_ascii=False, indent=2)


def storage_state_path(platform):
    p = os.path.join(DATA_DIR, f"{platform}_state.json")
    return p if os.path.exists(p) else None


def _stable_hash(s):
    """실행마다 값이 바뀌는 파이썬 내장 hash() 대신, 항상 같은 입력에 같은 값을 내는 해시.
    (내장 hash()는 프로세스마다 시드가 랜덤이라 매 실행 GitHub Actions job마다 값이 달라져서
    같은 게시물인데도 seen.json과 매번 다르게 매칭돼 계속 "새 글"로 잘못 잡히는 버그가 있었음)"""
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:16]


# ---------- 수집 ----------

def fetch_twitter_posts(page, handle, limit=5):
    """x.com/{handle} 최신 트윗을 읽어옵니다.

    X가 가끔 실제 article 요소를 "화면에 보이는" 상태로 그리기 전에
    검색엔진/봇용으로 보이는 정적 텍스트만 먼저 내려주는 경우가 있어,
    1) 우선 article이 DOM에 "붙기만"(attached) 해도 통과시켜 최대한 실제 요소를 노려보고,
    2) 그래도 실패하면 body 텍스트를 휴리스틱으로 파싱하는 최후 수단을 쓴다.
    2번 경로는 실제 트윗 ID를 알 수 없어 링크가 프로필 주소로 대체된다.
    """
    posts = []
    try:
        page.goto(f"https://x.com/{handle}", timeout=30000)
    except Exception as e:
        print(f"  [경고] @{handle} 트위터 페이지 이동 실패: {e}")
        return posts

    try:
        page.wait_for_selector('article[data-testid="tweet"]', timeout=15000, state="attached")
        articles = page.locator('article[data-testid="tweet"]').all()[:limit]
        if articles:
            return _extract_tweets_from_articles(articles, handle)
    except Exception as e:
        print(f"  [경고] @{handle} article 렌더링 대기 실패, 텍스트 파싱으로 대체 시도: {e}")

    print(f"  [디버그] 이동 후 URL: {page.url}")
    print(f"  [디버그] 페이지 제목: {page.title()}")
    body_text = page.locator("body").inner_text()
    print(f"  [디버그] 본문 앞부분: {body_text[:300]}")
    return _parse_tweets_from_text(body_text, handle, limit)


def _extract_tweets_from_articles(articles, handle):
    posts = []
    for a in articles:
        try:
            link = a.locator('a:has(time)').first
            href = link.get_attribute("href")
            post_id = href.split("/")[-1] if href else None
            time_el = a.locator("time").first
            posted_at = time_el.get_attribute("datetime")
            text_el = a.locator('[data-testid="tweetText"]').first
            text = text_el.inner_text() if text_el.count() else ""
            if post_id:
                posts.append({
                    "id": f"tw_{handle}_{post_id}",
                    "text": text.strip(),
                    "posted_at": posted_at,
                    "url": f"https://x.com/{handle}/status/{post_id}",
                })
        except Exception:
            continue
    return posts


# 트윗 작성 시각으로 흔히 나오는 상대/절대 시각 표기 패턴("18h", "3분", "7월 17일" 등)
_TW_TIME_PATTERN = re.compile(
    r'^(\d+(초|분|시간|일|주|개월|년|h|m|s)|\d{1,2}월\s*\d{1,2}일|\d{4}년\s*\d{1,2}월\s*\d{1,2}일)$'
)
# 좋아요/리트윗/조회수 등 참여수 숫자 라인("372", "1.5천", "3.6만" 등)
_TW_STAT_PATTERN = re.compile(r'^[\d,.\s]+(천|만)?$')


def _parse_tweets_from_text(body_text, handle, limit=5):
    """article 요소를 못 찾을 때 body 텍스트에서 트윗을 휴리스틱으로 파싱하는 최후 수단.
    실제 트윗 ID/링크를 알 수 없어 url은 프로필 주소로 대체된다."""
    lines = [l.strip() for l in body_text.split("\n") if l.strip()]
    posts = []
    marker = f"@{handle}"
    i = 0
    while i < len(lines) and len(posts) < limit:
        if lines[i] == marker and i + 1 < len(lines) and _TW_TIME_PATTERN.match(lines[i + 1]):
            posted_label = lines[i + 1]
            j = i + 2
            text_lines = []
            while j < len(lines):
                line = lines[j]
                if line == marker or _TW_STAT_PATTERN.match(line):
                    break
                text_lines.append(line)
                j += 1
            text = " ".join(text_lines).strip()
            if text:
                post_key = f"{posted_label}_{_stable_hash(text)}"
                posts.append({
                    "id": f"tw_{handle}_{post_key}",
                    "text": text[:500],
                    "posted_at": None,
                    "url": f"https://x.com/{handle}",
                })
            i = j
        else:
            i += 1
    return posts


def fetch_facebook_posts(page, page_name, limit=5):
    """facebook.com/{page_name} 최신 게시물을 읽어옵니다.
    로그인 세션이 있으면 mbasic이 자동으로 일반 facebook.com으로 리다이렉트되는 경우가 있어,
    role=article(ARIA) 선택자를 우선 사용하고 여러 후보 선택자를 순서대로 시도합니다.
    """
    posts = []
    try:
        page.goto(f"https://www.facebook.com/{page_name}", timeout=30000)
        page.wait_for_timeout(2000)
    except Exception as e:
        print(f"  [경고] {page_name} 페이스북 페이지 로드 실패: {e}")
        return posts

    print(f"  [디버그] 이동 후 URL: {page.url}")
    print(f"  [디버그] 페이지 제목: {page.title()}")

    articles = []
    for selector in ['[role="article"]', "article", "div[data-ft]"]:
        found = page.locator(selector).all()
        if found:
            print(f"  [디버그] 선택자 '{selector}'로 {len(found)}개 발견")
            articles = found[:limit]
            break

    print(f"  [디버그] 감지된 게시물 블록 수: {len(articles)}")
    if not articles:
        body_snippet = page.locator("body").inner_text()[:500]
        print(f"  [디버그] 본문 앞부분: {body_snippet}")

    for idx, a in enumerate(articles):
        try:
            text = extract_post_body_text(a)
            post_url = extract_post_permalink(a, page_name)
            dedup_key = _fb_dedup_key(post_url, page_name, idx)
            if text:
                posts.append({
                    "id": f"fb_{page_name}_{_stable_hash(dedup_key)}",
                    "text": text[:300],
                    "posted_at": None,
                    "url": post_url,
                })
        except Exception:
            continue
    return posts


_FB_PERMALINK_PATTERNS = ("/posts/", "permalink.php", "story_fbid", "/videos/", "/photo", "/reel/", "/watch/")

# 실제 게시물 permalink 안에서 "이 게시물"을 가리키는 고유 ID 부분만 뽑는 패턴.
# (story_fbid=123..., /posts/pfbid0Abc..., /videos/456... 등)
_FB_ID_PATTERN = re.compile(r'(?:story_fbid=|/posts/|/videos/|/reel/|/photo(?:\.php)?/?(?:fbid=)?)([A-Za-z0-9_-]+)')


def _fb_dedup_key(post_url, page_name, idx):
    """중복 판정용 안정 키. permalink에 남아있는 나머지 쿼리스트링(추적 파라미터 등)은
    스크랩할 때마다 값이 미묘하게 바뀔 수 있어, URL 전체 대신 게시물 고유 ID만 뽑아서 쓴다.
    진짜 permalink을 못 찾아 프로필 홈 주소로 대체된 경우엔 기존처럼 idx+날짜로 대체한다."""
    fallback_url = f"https://www.facebook.com/{page_name}"
    if post_url and post_url != fallback_url:
        m = _FB_ID_PATTERN.search(post_url)
        if m:
            return m.group(1)
        return post_url.split("?")[0]
    return f"{page_name}_{idx}_{datetime.now().date()}"


def extract_post_permalink(article_locator, page_name):
    """article 안의 링크들 중 실제 게시물 permalink으로 보이는 것을 찾는다.
    못 찾으면 계정 홈 주소로 대체한다."""
    fallback_url = f"https://www.facebook.com/{page_name}"
    try:
        links = article_locator.locator("a").all()
    except Exception:
        return fallback_url

    for link in links:
        try:
            href = link.get_attribute("href")
        except Exception:
            continue
        if not href:
            continue
        if any(p in href for p in _FB_PERMALINK_PATTERNS):
            if href.startswith("/"):
                href = "https://www.facebook.com" + href
            # 페이스북 추적 파라미터(__cft__, __tn__ 등) 제거
            href = href.split("&__cft__")[0].split("?__cft__")[0]
            href = href.split("&__tn__")[0]
            return href
    return fallback_url


# 게시물 텍스트에 섞여 들어오는 UI 요소(좋아요/댓글/공유 수, 버튼 라벨 등)를 걸러내는 패턴
_FB_NOISE_PATTERNS = (
    "좋아요", "댓글", "공유", "답글", "모든 공감", "전체보기", "더 보기", "더보기",
    "팔로우", "친구 추가", "메시지 보내기", "관련 콘텐츠",
)

# 긴 게시물을 접어서 보여줄 때 붙는 "더 보기" 류 버튼 라벨.
# 이걸 클릭해서 펼치지 않으면 본문이 미리보기 몇 글자 + 이 라벨로만 잘려서 스크랩된다.
_FB_EXPAND_LABELS = ("더 보기", "더보기", "See more", "See More")


def _expand_post_text(article_locator):
    """게시물이 '더 보기'로 접혀 있으면 클릭해서 전체 본문을 펼친다."""
    for selector in ['div[role="button"]', 'span[role="button"]']:
        try:
            buttons = article_locator.locator(selector).all()
        except Exception:
            continue
        for b in buttons:
            try:
                label = b.inner_text().strip()
            except Exception:
                continue
            if label in _FB_EXPAND_LABELS:
                try:
                    b.click(timeout=2000)
                    article_locator.page.wait_for_timeout(500)
                except Exception:
                    pass
                return


def extract_post_body_text(article_locator):
    """게시물 블록(article) 안에서 좋아요/댓글 수 같은 UI 텍스트를 뺀 실제 본문만 추출한다."""
    _expand_post_text(article_locator)

    # 1순위: 페이스북이 광고/게시물 본문에 붙이는 표준 마커
    for selector in ['[data-ad-preview="message"]', '[data-ad-comet-preview="message"]']:
        body_el = article_locator.locator(selector).first
        if body_el.count():
            body_text = body_el.inner_text().strip()
            if body_text:
                return body_text

    # 2순위: article 전체 텍스트에서 UI/숫자성 잡음 줄을 제거하는 휴리스틱
    raw = article_locator.inner_text().strip()
    lines = [line.strip() for line in raw.split("\n") if line.strip()]
    kept = []
    for line in lines:
        if any(noise in line for noise in _FB_NOISE_PATTERNS):
            continue
        if line.replace(",", "").replace(".", "").isdigit():
            continue
        if len(line) <= 1:
            continue
        kept.append(line)
    return " ".join(kept).strip()


# ---------- AI 요약 ----------

# 2026년 구글이 API 키를 "인증(auth) 키"(AQ.로 시작)로 전환하면서,
# 인증 방식도 URL의 ?key= 파라미터가 아니라 x-goog-api-key 요청 헤더로 바뀌었다.
GEMINI_MODEL = "gemini-3.5-flash"


def classify_and_summarize(text):
    """게시글이 AI(인공지능)/ICT 관련 내용인지 판별하고, 관련이 있으면 주제 요약도 함께 만든다.
    한 번의 Gemini 호출로 판별+요약을 같이 처리해서 API 호출 수를 늘리지 않는다.

    반환값: (is_relevant, summary_text)
    - GEMINI_API_KEY가 없거나 호출이 계속 실패하면, 걸러내지 못하고 원문 일부로 통과시킨다
      (필터링 오류로 진짜 관련 게시물을 놓치는 것보다는, 일단 보여주는 쪽이 안전하다고 판단).
    """
    fallback = (text or "").strip()[:300]
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("  [경고] GEMINI_API_KEY 환경변수가 비어있음 - AI/ICT 관련도 판별 없이 그대로 통과")
        return True, fallback
    if not text or not text.strip():
        return False, ""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    prompt = (
        "다음은 정치인/공직자의 SNS 게시글이야. 아래 두 가지를 판단해서 JSON 형식으로만 답해줘. "
        "다른 설명이나 코드블록 없이 JSON 객체 하나만 출력해.\n\n"
        "1) relevant: 이 글이 AI(인공지능)·ICT(정보통신기술) 정책/산업/기술/규제와 직접 관련이 있으면 true, "
        "일반 정치·의전·지역구 활동·무관한 주제면 false.\n"
        "2) summary: relevant가 true일 때만 작성. 이 글이 '무엇에 대한' 글인지 한국어 한 문장으로 압축해줘. "
        "'~했습니다/~합니다' 식으로 내용을 그대로 풀어 쓰지 말고, 핵심 주제·대상·쟁점이 드러나는 "
        "제목/헤드라인 톤으로. 예: '반도체 산업 지원 확대를 강조하는 발언'. relevant가 false면 빈 문자열.\n\n"
        '형식: {"relevant": true, "summary": "..."}\n\n'
        f"게시글:\n{text[:1500]}"
    )
    payload = {"contents": [{"parts": [{"text": prompt}]}]}

    # 무료 티어 분당 요청 한도(429)나 일시적 서버 과부하(503)에 대비해 짧게 재시도한다.
    max_attempts = 4
    backoff_seconds = [3, 8, 20]
    for attempt in range(max_attempts):
        try:
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "x-goog-api-key": api_key,
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            raw = data["candidates"][0]["content"]["parts"][0]["text"].strip()
            # 가끔 ```json ... ``` 코드블록으로 감싸서 나오는 경우가 있어 걷어낸다.
            if raw.startswith("```"):
                raw = raw.strip("`")
                if raw.lower().startswith("json"):
                    raw = raw[4:]
                raw = raw.strip()
            parsed = json.loads(raw)
            is_relevant = bool(parsed.get("relevant"))
            summary = (parsed.get("summary") or "").strip()
            if is_relevant and not summary:
                summary = fallback
            return is_relevant, summary
        except urllib.error.HTTPError as e:
            if e.code in (429, 503) and attempt < max_attempts - 1:
                wait = backoff_seconds[attempt]
                print(f"  [정보] Gemini {e.code} 응답, {wait}초 대기 후 재시도 ({attempt + 1}/{max_attempts})")
                time.sleep(wait)
                continue
            print(f"  [경고] AI 관련도 판별/요약 실패, 걸러내지 않고 원문 일부로 통과: {e}")
            return True, fallback
        except Exception as e:
            print(f"  [경고] AI 관련도 판별/요약 실패, 걸러내지 않고 원문 일부로 통과: {e}")
            return True, fallback
    return True, fallback


# ---------- 이메일 ----------

def build_excel_from_csv():
    """new_items.csv(누적, 최신순) 전체를 계정명/플랫폼/게시시각/요약/본문 링크 5개 컬럼 엑셀로 만든다."""
    if not os.path.exists(CSV_PATH):
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "SNS 모니터링"
    ws.append(["계정명", "플랫폼", "게시시각", "요약", "본문 링크"])

    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ws.append([
                row.get("인물", ""),
                row.get("플랫폼", ""),
                row.get("게시시각표기", ""),
                row.get("내용요약", ""),
                row.get("링크", ""),
            ])

    for col_idx, width in enumerate([20, 20, 18, 60, 40], start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    os.makedirs(DATA_DIR, exist_ok=True)
    wb.save(EXCEL_PATH)
    return EXCEL_PATH


def build_html_from_csv():
    """new_items.csv(누적, 최신순) 전체를, 엑셀보다 바로 읽기 편하도록 스타일 입힌
    HTML 표로 만든다. GitHub Pages가 이 파일을 그대로 고정 URL로 서빙한다."""
    if not os.path.exists(CSV_PATH):
        return None

    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = list(csv.DictReader(f))

    rows_html = []
    for row in reader:
        person = html.escape(row.get("인물", ""))
        platform = html.escape(row.get("플랫폼", ""))
        posted_at = html.escape(row.get("게시시각표기", "") or "확인불가")
        summary = html.escape(row.get("내용요약", "")).replace("\n", "<br>")
        link = html.escape(row.get("링크", ""), quote=True)
        rows_html.append(
            "<tr>"
            f"<td class='c-person'>{person}</td>"
            f"<td class='c-platform'>{platform}</td>"
            f"<td class='c-time'>{posted_at}</td>"
            f"<td class='c-summary'>{summary}</td>"
            f"<td class='c-link'><a href=\"{link}\" target=\"_blank\">원문 보기</a></td>"
            "</tr>"
        )

    doc = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>SNS 모니터링</title>
<style>
  body {{ font-family: -apple-system, "Malgun Gothic", "Apple SD Gothic Neo", sans-serif;
          background: #f4f5f7; margin: 0; padding: 24px; color: #1f2328; }}
  h1 {{ font-size: 18px; margin: 0 0 16px; }}
  table {{ width: 100%; border-collapse: collapse; background: #fff;
           box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-radius: 8px; overflow: hidden; }}
  th, td {{ padding: 10px 12px; text-align: left; border-bottom: 1px solid #e5e7eb;
            vertical-align: top; font-size: 14px; }}
  th {{ background: #1f2937; color: #fff; font-weight: 600; position: sticky; top: 0; }}
  tr:nth-child(even) {{ background: #fafafa; }}
  tr:hover {{ background: #eef2ff; }}
  .c-person {{ white-space: nowrap; font-weight: 600; }}
  .c-platform {{ white-space: nowrap; color: #555; }}
  .c-time {{ white-space: nowrap; color: #777; font-size: 13px; }}
  .c-summary {{ max-width: 480px; }}
  .c-link a {{ color: #2563eb; text-decoration: none; white-space: nowrap; }}
  .c-link a:hover {{ text-decoration: underline; }}
</style>
</head>
<body>
  <h1>SNS 모니터링 ({len(reader)}건, 최신순)</h1>
  <table>
    <thead>
      <tr><th>계정명</th><th>플랫폼</th><th>게시시각</th><th>요약</th><th>본문 링크</th></tr>
    </thead>
    <tbody>
      {''.join(rows_html)}
    </tbody>
  </table>
</body>
</html>"""

    os.makedirs(DOCS_DIR, exist_ok=True)
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(doc)
    return HTML_PATH


def send_email(env, subject, body, attachment_paths=None):
    context = ssl.create_default_context()
    attachment_paths = [p for p in (attachment_paths or []) if p and os.path.exists(p)]

    if attachment_paths:
        msg = MIMEMultipart()
        msg.attach(MIMEText(body))
        for attachment_path in attachment_paths:
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            filename = os.path.basename(attachment_path)
            part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
            msg.attach(part)
    else:
        msg = MIMEText(body)

    msg["Subject"] = subject
    msg["From"] = env["SMTP_FROM"]
    msg["To"] = env["EMAIL_TO"]

    with smtplib.SMTP(env["SMTP_HOST"], int(env["SMTP_PORT"])) as server:
        server.starttls(context=context)
        server.login(env["SMTP_USERNAME"], env["SMTP_PASSWORD"])
        server.sendmail(env["SMTP_FROM"], [env["EMAIL_TO"]], msg.as_string())


def format_report(new_items):
    lines = [f"# SNS 모니터링 새 게시물 ({datetime.now().strftime('%Y-%m-%d %H:%M')})", ""]
    for item in new_items:
        lines.append(f"## {item['person']} ({item['platform']})")
        lines.append(f"- 게시시각: {item['post'].get('posted_at') or '확인불가'}")
        lines.append(f"- 내용: {item['post']['text'][:300]}")
        lines.append(f"- 링크: {item['post']['url']}")
        lines.append("")
    return "\n".join(lines)


def sort_newest_first(new_items):
    """posted_at(ISO 문자열)이 있으면 그 기준, 없으면 원래 순서 유지하며 뒤로 보냄."""
    def key(item):
        posted_at = item["post"].get("posted_at")
        return posted_at or ""  # ISO 문자열은 그대로 비교해도 최신이 더 크게 나옴
    return sorted(new_items, key=key, reverse=True)


def csv_escape(value):
    value = "" if value is None else str(value)
    if any(c in value for c in [",", "\n", '"']):
        value = '"' + value.replace('"', '""') + '"'
    return value


def prepend_csv_rows(new_items, run_time_str):
    """new_items.csv 헤더 바로 다음(파일 최상단)에 이번에 새로 감지된 행들을 끼워 넣는다."""
    os.makedirs(DATA_DIR, exist_ok=True)

    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", encoding="utf-8") as f:
            existing_lines = f.readlines()
        existing_rows = existing_lines[1:] if existing_lines else []
    else:
        existing_rows = []

    new_rows = []
    for item in new_items:
        row = [
            run_time_str,
            item["person"],
            item["platform"],
            item["post"].get("posted_at") or "확인불가",
            item["post"]["text"][:300],
            item["post"]["url"],
        ]
        new_rows.append(",".join(csv_escape(v) for v in row) + "\n")

    with open(CSV_PATH, "w", encoding="utf-8") as f:
        f.write(CSV_HEADER)
        f.writelines(new_rows)
        f.writelines(existing_rows)


# ---------- 메인 ----------

def main():
    env = load_env(ENV_PATH)
    persons = load_sources()
    seen = load_seen()

    new_items = []

    # 헤드리스 브라우저는 X 등에서 봇으로 더 쉽게 감지되는 경향이 있어,
    # HEADLESS=false 환경변수(GitHub Actions에서는 xvfb와 함께 사용)로 일반 브라우저처럼 띄운다.
    run_headless = os.environ.get("HEADLESS", "true").lower() != "false"
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=run_headless,
            args=["--disable-blink-features=AutomationControlled"],
        )

        for person in persons:
            for handle in person.get("twitter", []):
                state = storage_state_path("twitter")
                context = browser.new_context(**({"storage_state": state} if state else {}), **CONTEXT_ARGS)
                page = context.new_page()
                print(f"확인 중: {person['name']} - X @{handle}")
                posts = fetch_twitter_posts(page, handle)
                seen_key = f"tw_{handle}"
                seen_ids = set(seen.get(seen_key, []))
                for post in posts:
                    if post["id"] not in seen_ids:
                        new_items.append({"person": person["name"], "platform": f"X(@{handle})", "post": post})
                        seen_ids.add(post["id"])
                seen[seen_key] = list(seen_ids)
                context.close()

            for fb_page in person.get("facebook", []):
                state = storage_state_path("facebook")
                context = browser.new_context(**({"storage_state": state} if state else {}), **CONTEXT_ARGS)
                page = context.new_page()
                print(f"확인 중: {person['name']} - Facebook {fb_page}")
                posts = fetch_facebook_posts(page, fb_page)
                seen_key = f"fb_{fb_page}"
                seen_ids = set(seen.get(seen_key, []))
                for post in posts:
                    if post["id"] not in seen_ids:
                        new_items.append({"person": person["name"], "platform": f"Facebook({fb_page})", "post": post})
                        seen_ids.add(post["id"])
                seen[seen_key] = list(seen_ids)
                context.close()

        browser.close()

    save_seen(seen)

    if not new_items:
        print("새 게시물 없음, 리포트/CSV/메일 모두 건너뜀")
        return

    new_items = sort_newest_first(new_items)

    # AI/ICT 관련 게시물만 남기고, 무관한 건 리포트/CSV/엑셀/메일에서 제외한다.
    # (dedup용 seen 기록은 위에서 이미 저장됐으므로, 여기서 걸러져도 다음 실행 때 다시 안 잡힌다.)
    # SKIP_RELEVANCE_FILTER=true면 필터를 끄고 새 글이면 무조건 통과시킨다.
    # (스크래핑/메일 파이프라인 자체가 실제로 잘 도는지 테스트할 때, "지금 AI/ICT 글이
    #  없어서 안 오는 건지" vs "스크래핑이 고장나서 안 오는 건지" 구분하기 위한 용도)
    skip_filter = os.environ.get("SKIP_RELEVANCE_FILTER", "false").lower() == "true"
    if skip_filter:
        print("  [정보] SKIP_RELEVANCE_FILTER=true - AI/ICT 필터 끄고 새 글 전부 통과")

    filtered_items = []
    for idx, item in enumerate(new_items):
        is_relevant, summary = classify_and_summarize(item["post"]["text"])
        if skip_filter:
            is_relevant = True
            if not summary:
                summary = item["post"]["text"][:300]
        if is_relevant:
            item["post"]["text"] = summary
            filtered_items.append(item)
        else:
            print(f"  [정보] AI/ICT 무관 게시물 제외: {item['person']} ({item['platform']})")
        # 무료 티어 분당 요청 한도(429)를 애초에 안 넘도록 호출 사이 텀을 둔다.
        if idx < len(new_items) - 1:
            time.sleep(5)
    new_items = filtered_items

    if not new_items:
        print("AI/ICT 관련 새 게시물 없음, 리포트/CSV/메일 모두 건너뜀")
        return

    run_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    os.makedirs(REPORTS_DIR, exist_ok=True)
    report_path = os.path.join(REPORTS_DIR, f"{datetime.now().strftime('%Y%m%d_%H%M')}.md")
    report_text = format_report(new_items)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_text)

    prepend_csv_rows(new_items, run_time_str)
    excel_path = build_excel_from_csv()
    html_path = build_html_from_csv()

    # HTML은 파일로 첨부하는 대신, GitHub Pages 고정 링크로 안내한다.
    # (링크는 항상 동일하고, 페이지 내용만 매번 최신으로 갱신됨)
    pages_url = env.get("PAGES_URL", "").strip()
    if pages_url:
        body_text = (
            report_text
            + f"\n\n---\n전체 최신 목록 보기(항상 최신으로 갱신됨): {pages_url}\n"
        )
    else:
        body_text = report_text
        print("  [경고] PAGES_URL이 설정 안 됨 - 메일 본문에 고정 링크를 못 넣었음")

    subject = f"[SNS 모니터링] 새 게시물 {len(new_items)}건 업데이트"
    send_email(env, subject, body_text, attachment_paths=[excel_path])
    print(f"메일 발송 완료: {len(new_items)}건 (엑셀 첨부: {excel_path}, 페이지: {html_path})")


if __name__ == "__main__":
    main()
