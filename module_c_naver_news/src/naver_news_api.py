import html
import requests
from email.utils import parsedate_to_datetime


NAVER_NEWS_URL = "https://openapi.naver.com/v1/search/news.json"


def clean_html_text(text: str) -> str:
    if not text:
        return ""
    text = html.unescape(text)
    text = text.replace("<b>", "").replace("</b>", "")
    return text.strip()


def parse_pubdate(pub_date: str) -> str:
    try:
        dt = parsedate_to_datetime(pub_date)
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return pub_date


def fetch_news_by_query(
    keyword: str,
    search_query: str,
    client_id: str,
    client_secret: str,
    display: int = 10,
    start: int = 1,
    sort: str = "date",
    timeout: int = 10,
) -> list[dict]:
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }

    params = {
        "query": search_query,
        "display": display,
        "start": start,
        "sort": sort,
    }

    response = requests.get(
        NAVER_NEWS_URL,
        headers=headers,
        params=params,
        timeout=timeout,
    )
    response.raise_for_status()

    data = response.json()
    items = data.get("items", [])

    rows = []
    for idx, item in enumerate(items, start=1):
        title = clean_html_text(item.get("title", ""))
        description = clean_html_text(item.get("description", ""))
        pub_date = parse_pubdate(item.get("pubDate", ""))
        originallink = item.get("originallink", "")
        naver_link = item.get("link", "")

        final_link = originallink if originallink else naver_link

        rows.append({
            "키워드": keyword,
            "검색어": search_query,
            "순위": idx,
            "기사제목": title,
            "날짜": pub_date,
            "요약": description,
            "기사링크": final_link,
        })

    return rows