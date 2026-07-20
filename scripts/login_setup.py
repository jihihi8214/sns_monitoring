"""
X(트위터), 페이스북에 '직접' 로그인해서 세션(쿠키)을 저장하는 1회성 스크립트.

- 헤드리스가 아니라 실제 브라우저 창이 뜹니다.
- 본인이 직접 아이디/비밀번호를 입력하고 로그인하세요 (이 스크립트는 자동 로그인하지 않습니다).
- 로그인 후 터미널에서 Enter를 누르면 세션이 data/ 폴더에 저장되고,
  이후 monitor.py 실행 시 로그인된 상태로 페이지를 읽어옵니다.

실행:
    python3 scripts/login_setup.py twitter
    python3 scripts/login_setup.py facebook
"""

import sys
import os
from playwright.sync_api import sync_playwright

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE, "data")

TARGETS = {
    "twitter": "https://x.com/login",
    "facebook": "https://www.facebook.com/login",
}


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in TARGETS:
        print("사용법: python3 scripts/login_setup.py [twitter|facebook]")
        sys.exit(1)

    platform = sys.argv[1]
    os.makedirs(DATA_DIR, exist_ok=True)
    storage_path = os.path.join(DATA_DIR, f"{platform}_state.json")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto(TARGETS[platform])

        input(f"\n{platform} 로그인을 마친 뒤 이 터미널에서 Enter를 누르세요...\n")

        context.storage_state(path=storage_path)
        browser.close()

    print(f"저장 완료: {storage_path}")


if __name__ == "__main__":
    main()
