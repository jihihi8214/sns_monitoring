import os
import smtplib
import ssl
from email.mime.text import MIMEText

def load_env(path):
    env = {}
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
env = load_env(os.path.join(BASE, ".env"))

msg = MIMEText("SNS 모니터링 시스템 SMTP 발송 테스트입니다. 이 메일을 받으셨다면 설정이 정상입니다.")
msg["Subject"] = "[테스트] SNS 모니터링 메일 발송 확인"
msg["From"] = env["SMTP_FROM"]
msg["To"] = env["EMAIL_TO"]

try:
    context = ssl.create_default_context()
    with smtplib.SMTP(env["SMTP_HOST"], int(env["SMTP_PORT"])) as server:
        server.starttls(context=context)
        server.login(env["SMTP_USERNAME"], env["SMTP_PASSWORD"])
        server.sendmail(env["SMTP_FROM"], [env["EMAIL_TO"]], msg.as_string())
    print("SUCCESS: 메일 발송 성공")
except Exception as e:
    print(f"FAIL: {type(e).__name__}: {e}")
