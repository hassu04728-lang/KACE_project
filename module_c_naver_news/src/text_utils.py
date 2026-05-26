import re
from collections import Counter


STOPWORDS = {
    "있다", "없다", "대한", "위한", "통해", "관련", "기자", "뉴스", "기사",
    "이번", "최근", "이날", "지난", "오늘", "오전", "오후", "시장", "산업",
    "분야", "기업", "국내", "글로벌", "정부", "중심", "확대", "추진", "발표",
    "지원", "전망", "증가", "감소", "변화", "진행", "기반", "정도", "경우",
    "채용", "공채", "인턴"  # 채용 검색어 확장 때문에 너무 많이 잡히는 단어 제거
}


def tokenize(text: str) -> list[str]:
    if not text:
        return []

    text = re.sub(r"[^0-9A-Za-z가-힣\s]", " ", text)
    words = re.findall(r"[0-9A-Za-z가-힣]{2,}", text)

    result = []
    for word in words:
        if word in STOPWORDS:
            continue
        if word.isdigit():
            continue
        result.append(word)
    return result


def build_trend_paragraph(keyword: str, search_query: str, news_rows: list[dict]) -> str:
    if not news_rows:
        return f"{keyword} 관련 기사가 수집되지 않아 최신 동향을 정리하지 못했다."

    combined = " ".join(
        f"{row.get('기사제목', '')} {row.get('요약', '')}"
        for row in news_rows
    )

    tokens = tokenize(combined)
    counter = Counter(tokens)
    top_terms = [word for word, _ in counter.most_common(6)]

    if not top_terms:
        return (
            f"{keyword} 관련 상위 기사들을 종합하면, 최근에는 해당 분야에서 정책 변화, "
            f"시장 흐름, 기술 개발 및 기관·기업의 움직임이 함께 나타나고 있다. "
            f"단기 이슈보다는 향후 방향성과 실제 적용 가능성을 보여주는 기사들이 주로 노출되는 경향이 있다."
        )

    term_text = ", ".join(top_terms[:5])

    return (
        f"{keyword} 관련 상위 기사들을 종합하면 최근 흐름은 {term_text} 등을 중심으로 형성되고 있다. "
        f"전반적으로는 단순 사건성 보도보다 기술 개발, 투자·사업 확대, 제도 변화, 기관 및 기업의 전략 움직임이 "
        f"함께 나타나는 모습이며, 이를 통해 {keyword} 분야가 현재 어떤 주제에 관심이 집중되고 있는지 확인할 수 있다. "
        f"이번 요약은 검색어 '{search_query}'로 수집한 기사 제목과 요약문을 바탕으로 정리한 결과다."
    )