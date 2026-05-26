from pathlib import Path
import pandas as pd


SEARCH_QUERY_MAP = {
    "반도체": "반도체",
    "철강": "철강 OR 금속",
    "채용": "채용 OR 공채 OR 인턴",
    "세라믹": "세라믹",
    "고분자": "고분자",
    "디스플레이": "디스플레이",
    "고려대학교": "고려대학교",
}


def load_active_keywords(config_path: str, sheet_name: str = "설정") -> list[dict]:
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"설정 파일이 없음: {config_path}")

    df = pd.read_excel(path, sheet_name=sheet_name)

    required_cols = {"키워드", "사용여부"}
    if not required_cols.issubset(df.columns):
        raise ValueError("설정 시트에 '키워드', '사용여부' 열이 있어야 함")

    df["키워드"] = df["키워드"].astype(str).str.strip()
    df["사용여부"] = df["사용여부"].astype(str).str.strip().str.upper()

    active = []
    for _, row in df.iterrows():
        if row["사용여부"] == "Y":
            keyword = row["키워드"]
            search_query = SEARCH_QUERY_MAP.get(keyword, keyword)
            active.append({
                "키워드": keyword,
                "검색어": search_query
            })

    if not active:
        raise ValueError("사용여부가 Y인 키워드가 없음")

    return active