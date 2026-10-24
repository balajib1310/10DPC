
import requests
from bs4 import BeautifulSoup

URL = "https://www.bbc.com/news"

headers = {
    "User-Agent": "Mozilla/5.0"
}

try:
    response = requests.get(URL, headers=headers, timeout=10)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    headlines = soup.select(
        "h2.IndexCardHeading-styles__CardHeadlineStyled-sc-c7d910a6-3"
    )

    unique_headlines = []
    seen = set()

    for item in headlines:
        title = item.get_text(" ", strip=True)

        if not title or title in seen:
            continue

        seen.add(title)
        unique_headlines.append(title)

        if len(unique_headlines) == 5:
            break

    if len(unique_headlines) < 5:
        print(f"Only {len(unique_headlines)} headlines found.")

    for index, title in enumerate(unique_headlines, start=1):
        print(f"{index}. {title}")

except requests.RequestException as error:
    print(f"Failed to fetch news: {error}")
