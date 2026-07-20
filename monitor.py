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

    # mbasic 구조는 자주 바뀔 수 있어 텍스트 블록 단위로 느슨하게 추출
    articles = page.locator("article").all()[:limit]
    if not articles:
        # article 태그가 없을 경우 fallback: 본문 텍스트 블록 시도
        articles = page.locator("div[data-ft]").all()[:limit]

    for idx, a in enumerate(articles):
        try:
            text = a.inner_text().strip()
            link_el = a.locator("a").first
            href = link_el.get_attribute("href") if link_el.count() else None
            post_id = href if href else f"{page_name}_{idx}_{datetime.now().date()}"
            if text:
                posts.append({
                    "id": f"fb_{page_name}_{hash(post_id)}",
                    "text": text[:300],
                    "posted_at": None,
                    "url": f"https://www.facebook.com/{page_name}",
                })
        except Exception:
            continue
    return posts


# ---------- 이메일 ----------

def send_email(env, subject, body):
    context = ssl.create_default_context()
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

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        for person in persons:
            for handle in person.get("twitter", []):
                state = storage_state_path("twitter")
                context = browser.new_context(storage_state=state) if state else browser.new_context()
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
                context = browser.new_context(storage_state=state) if state else browser.new_context()
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
    run_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    os.makedirs(REPORTS_DIR, exist_ok=True)
    report_path = os.path.join(REPORTS_DIR, f"{datetime.now().strftime('%Y%m%d_%H%M')}.md")
    report_text = format_report(new_items)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_text)

    prepend_csv_rows(new_items, run_time_str)

    subject = f"[SNS 모니터링] 새 게시물 {len(new_items)}건"
    send_email(env, subject, report_text)
    print(f"메일 발송 완료: {len(new_items)}건")


if __name__ == "__main__":
    main()
