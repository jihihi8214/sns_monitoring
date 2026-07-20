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

# 봇
