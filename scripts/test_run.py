"""
테스트용 스크립트: 실제 X/페이스북 스크래핑 없이,
더미 AI/ICT 게시물 하나로 monitor.py의 실제 파이프라인
(Gemini 관련도 판별+요약 -> 리포트/엑셀/HTML 생성 -> 메일 발송)이 잘 도는지 확인한다.

- seen.json(중복 판정 기록)은 절대 건드리지 않는다.
- new_items.csv / sns_monitoring.xlsx / sns_monitoring.html 같은 실제 누적 데이터도
  건드리지 않고, data/test_* 이름의 별도 파일에만 쓴다.
- 즉 이 스크립트를 몇 번을 돌려도 운영 데이터는 전혀 영향받지 않는다.

실행:
    python3 scripts/test_run.py
"""

import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import monitor as m

# 실제 누적 파일 대신 테스트 전용 파일 경로로 바꿔치기
m.CSV_PATH = os.path.join(m.DATA_DIR, "test_new_items.csv")
m.EXCEL_PATH = os.path.join(m.DATA_DIR, "test_sns_monitoring.xlsx")
m.HTML_PATH = os.path.join(m.DATA_DIR, "test_sns_monitoring.html")


def main():
    env = m.load_env(m.ENV_PATH)

    dummy_post = {
        "id": "test_dummy",
        "text": (
            "정부가 국가 AI 경쟁력 강화를 위해 2027년까지 프론티어 AI 모델 개발에 "
            "3조5000억원 규모의 예산을 투입하겠다고 밝혔다. 민간 기업과의 협력을 통해 "
            "글로벌 AI 3강 진입을 목표로 한다는 방침이다."
        ),
        "posted_at": datetime.now(timezone.utc).isoformat(),
        "url": "https://example.com/test-post",
    }
    new_items = [{
        "person": "[테스트] 배경훈",
        "platform": "X(@msitminister)",
        "post": dummy_post,
    }]

    is_relevant, summary = m.classify_and_summarize(dummy_post["text"])
    print(f"[테스트] Gemini 판별 결과: relevant={is_relevant}, summary={summary!r}")
    new_items[0]["post"]["text"] = summary if is_relevant else dummy_post["text"][:300]

    run_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    os.makedirs(m.REPORTS_DIR, exist_ok=True)
    report_path = os.path.join(m.REPORTS_DIR, f"TEST_{datetime.now().strftime('%Y%m%d_%H%M')}.md")
    report_text = m.format_report(new_items)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_text)

    m.prepend_csv_rows(new_items, run_time_str)
    excel_path = m.build_excel_from_csv()
    html_path = m.build_html_from_csv()

    subject = "[SNS 모니터링] 테스트 메일 (실제 게시물 아님)"
    m.send_email(env, subject, report_text, attachment_paths=[html_path, excel_path])
    print(f"[테스트] 메일 발송 완료 (첨부: {html_path}, {excel_path})")


if __name__ == "__main__":
    main()
