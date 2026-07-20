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
import csv
import json
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

try:
    import gspread
    from google.oauth2.service_account import Credentials as GoogleCredentials
except ImportError:
    gspread = None
    GoogleCredentials = None

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
REPORTS_DIR = os.path.join(BASE, "reports")
SEEN_PATH = os.path.join(DATA_DIR, "seen.json")
SOURCES_PATH = os.path.join(BASE, "sources.json")
ENV_PATH = os.path.join(BASE, ".env")
CSV_PATH = os.path.join(DATA_DIR, "new_items.csv")
CSV_HEADER = "확인시각,인물,플랫폼,게시시각표기,내용요약,링크\n"
EXCEL_PATH = os.path.join(DATA_DIR, "sns_monitoring.xlsx")

CONTEXT_ARGS = {
    "user_agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "viewport": {"width": 1366, "height": 900},
    "locale": "ko-KR",
}


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


def fetch_twitter_posts(page, handle, limit=5):
    posts = []
    try:
        page.goto(f"https://x.com/{handle}", timeout=30000)
        page.wait_for_selector('article[data-testid="tweet"]', timeout=15000)
    except Exception as e:
        print(f"  [경고] @{handle} 트위터 페이지 로드 실패: {e}")
        print(f"  [디버그] 이동 후 URL: {page.url}")
        print(f"  [디버그] 페이지 제목: {page.title()}")
        body_snippet = page.locator("body").inner_text()[:300]
        print(f"  [디버그] 본문 앞부분: {body_snippet}")
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
            post_id = post_url if post_url else f"{page_name}_{idx}_{datetime.now().date()}"
            if text:
                posts.append({
                    "id": f"fb_{page_name}_{hash(post_id)}",
                    "text": text[:300],
                    "posted_at": None,
                    "url": post_url,
                })
        except Exception:
            continue
    return posts


_FB_PERMALINK_PATTERNS = ("/posts/", "permalink.php", "story_fbid", "/videos/", "/photo", "/reel/", "/watch/")


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
            href = href.split("&__cft__")[0].split("?__cft__")[0]
            href = href.split("&__tn__")[0]
            return href
    return fallback_url


_FB_NOISE_PATTERNS = (
    "좋아요", "댓글", "공유", "답글", "모든 공감", "전체보기", "더 보기",
    "팔로우", "친구 추가", "메시지 보내기", "관련 콘텐츠",
)


def extract_post_body_text(article_locator):
    for selector in ['[data-ad-preview="message"]', '[data-ad-comet-preview="message"]']:
        body_el = article_locator.locator(selector).first
        if body_el.count():
            body_text = body_el.inner_text().strip()
            if body_text:
                return body_text

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


GEMINI_MODEL = "gemini-2.5-flash"


def summarize_text(text):
    fallback = (text or "").strip()[:300]
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key or not text or not text.strip():
        return fallback

    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent?key={api_key}"
    )
    prompt = (
        "다음은 정치인/공직자의 SNS 게시글이야. 핵심 내용만 한국어 1~2문장으로 짧게 요약해줘. "
        "원문을 그대로 인용하지 말고, 불필요하게 길게 늘이지 마. 요약문만 출력해.\n\n"
        f"게시글:\n{text[:1500]}"
    )
    payload = {"contents": [{"parts": [{"text": prompt}]}]}

    try:
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        summary = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        return summary if summary else fallback
    except Exception as e:
        print(f"  [경고] AI 요약 실패, 원문 일부로 대체: {e}")
        return fallback


def build_excel_from_csv():
    if not os.path.exists(CSV_PATH):
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "SNS 모니터링"
    ws.append(["계정명", "플랫폼", "요약", "본문 링크"])

    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ws.append([
                row.get("인물", ""),
                row.get("플랫폼", ""),
                row.get("내용요약", ""),
                row.get("링크", ""),
            ])

    for col_idx, width in enumerate([20, 20, 60, 40], start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    os.makedirs(DATA_DIR, exist_ok=True)
    wb.save(EXCEL_PATH)
    return EXCEL_PATH


def send_email(env, subject, body, attachment_path=None):
    context = ssl.create_default_context()

    if attachment_path and os.path.exists(attachment_path):
        msg = MIMEMultipart()
        msg.attach(MIMEText(body))
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


def upload_to_google_sheet(new_items):
    sa_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_PATH")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")

    if not sa_path or not sheet_id:
        print("  [정보] 구글 시트 설정 없음(GOOGLE_SERVICE_ACCOUNT_PATH/GOOGLE_SHEET_ID) - 건너뜀")
        return
    if gspread is None:
        print("  [경고] gspread 미설치 - 구글 시트 업로드 건너뜀")
        return

    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = GoogleCredentials.from_service_account_file(sa_path, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(sheet_id)
        ws = sh.sheet1

        header = ["계정명", "플랫폼", "요약", "본문 링크"]
        if ws.row_count == 0 or not ws.row_values(1):
            ws.update("A1", [header])

        rows = [
            [item["person"], item["platform"], item["post"]["text"][:300], item["post"]["url"]]
            for item in new_items
        ]
        ws.insert_rows(rows, row=2, value_input_option="RAW")
        print(f"  [정보] 구글 시트에 {len(rows)}행 추가 완료")
    except Exception as e:
        print(f"  [경고] 구글 시트 업로드 실패: {e}")


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
    def key(item):
        posted_at = item["post"].get("posted_at")
        return posted_at or ""
    return sorted(new_items, key=key, reverse=True)


def csv_escape(value):
    value = "" if value is None else str(value)
    if any(c in value for c in [",", "\n", '"']):
        value = '"' + value.replace('"', '""') + '"'
    return value


def prepend_csv_rows(new_items, run_time_str):
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


def main():
    env = load_env(ENV_PATH)
    persons = load_sources()
    seen = load_seen()

    new_items = []

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

    for item in new_items:
        item["post"]["text"] = summarize_text(item["post"]["text"])

    run_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    os.makedirs(REPORTS_DIR, exist_ok=True)
    report_path = os.path.join(REPORTS_DIR, f"{datetime.now().strftime('%Y%m%d_%H%M')}.md")
    report_text = format_report(new_items)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_text)

    prepend_csv_rows(new_items, run_time_str)
    upload_to_google_sheet(new_items)
    excel_path = build_excel_from_csv()

    subject = f"[SNS 모니터링] 새 게시물 {len(new_items)}건"
    send_email(env, subject, report_text, attachment_path=excel_path)
    print(f"메일 발송 완료: {len(new_items)}건 (엑셀 첨부: {excel_path})")


if __name__ == "__main__":
    main()
