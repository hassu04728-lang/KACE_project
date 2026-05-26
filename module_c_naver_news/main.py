import os
from dotenv import load_dotenv

from src.config_reader import load_active_keywords
from src.naver_news_api import fetch_news_by_query
from src.text_utils import build_trend_paragraph
from src.excel_writer import write_news_sheet, write_trend_sheet


CONFIG_PATH = "config/keyword_config.xlsx"


def main():
    load_dotenv()

    client_id = os.getenv("NAVER_CLIENT_ID")
    client_secret = os.getenv("NAVER_CLIENT_SECRET")

    if not client_id or not client_secret:
        raise ValueError(".env 파일에 NAVER_CLIENT_ID와 NAVER_CLIENT_SECRET을 입력해야 함")

    active_items = load_active_keywords(CONFIG_PATH)

    all_news_rows = []
    all_trend_rows = []

    print("선택된 키워드 목록")
    for item in active_items:
        print(f"- {item['키워드']} → 검색어: {item['검색어']}")

    for item in active_items:
        keyword = item["키워드"]
        search_query = item["검색어"]

        print(f"\n[수집 시작] {keyword}")

        try:
            news_rows = fetch_news_by_query(
                keyword=keyword,
                search_query=search_query,
                client_id=client_id,
                client_secret=client_secret,
                display=10,
                start=1,
                sort="date",
            )

            all_news_rows.extend(news_rows)

            trend_text = build_trend_paragraph(keyword, search_query, news_rows)
            all_trend_rows.append({
                "키워드": keyword,
                "검색어": search_query,
                "기사건수": len(news_rows),
                "종합 트렌드/최신 동향": trend_text,
            })

            print(f"[완료] {keyword}: {len(news_rows)}건 수집")

        except Exception as e:
            print(f"[오류] {keyword}: {e}")
            all_trend_rows.append({
                "키워드": keyword,
                "검색어": search_query,
                "기사건수": 0,
                "종합 트렌드/최신 동향": f"오류로 인해 수집 실패: {e}",
            })

    write_news_sheet(CONFIG_PATH, all_news_rows)
    write_trend_sheet(CONFIG_PATH, all_trend_rows)

    print(f"\n모든 작업 완료: {CONFIG_PATH}")
    print("엑셀 파일의 '뉴스 수집', '트렌드 요약' 시트를 확인하면 됨")


if __name__ == "__main__":
    main()