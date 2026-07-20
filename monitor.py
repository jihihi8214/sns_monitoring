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
import json
import smtplib
import ssl
from datetime import datetime
from email.mime.text import MIMEText

from playwright.sync_api import sync_playwright

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
REPORTS_DIR = os.path.join(BASE, "reports")
SEEN_PATH = os.path.join(DATA_DIR, "seen.json")
SOURCES_PATH = os.path.join(BASE, "sources.json")
ENV_PATH = os.path.join(BASE, ".env")
CSV_PATH = os.path.join(DATA_DIR, "new_items.csv")
CSV_HEADER = "확인시각,인물,플랫폼,게시시각표기,내용요약,링크\n"


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


# ---------- 수집 ----------

def fetch_twitter_posts(page, handle, limit=5):
    """x.com/{handle} 최신 트윗을 읽어옵니다."""
    posts = []
    try:
        page.goto(f"https://x.com/{handle}", timeout=30000)
        page.wait_for_selector('article[data-testid="tweet"]', timeout=15000)
    except Exception as e:
        print(f"  [경고] @{handle} 트위터 페이지 로드 실패: {e}")
        return posts

    articles = page.locator('article[data-testid="tweet"]').all()[:limit]
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


def fetch_facebook_posts(page, page_name, limit=5):
    """mbasic.facebook.com/{page_name} 최신 게시물을 읽어옵니다 (경량 HTML 버전)."""
    posts = []
    try:
        page.goto(f"https://mbasic.facebook.com/{page_name}", timeout=30000)
    except Exception as e:
        print(f"  [경고] {page_name} 페이스북 페이지 로드 실패: {e}")
        return posts

    print(f"  [디버그] 이동 후 URL: {page.url}")
    print(f"  [디버그] 페이지 제목: {page.title()}")

    # mbasic
