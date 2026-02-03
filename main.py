import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

URL = "https://in.bookmyshow.com/explore/events-bengaluru"
HEADERS = {"User-Agent": "Mozilla/5.0"}
FILE_NAME="events.xlsx"


def fetch_events():
    try:
        response = requests.get(
            URL,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Accept-Language": "en-IN,en;q=0.9",
                "Referer": "https://www.google.com"
            },
           timeout=10
        )

        print("Status Code:", response.status_code)

        if response.status_code != 200:
            raise Exception("Blocked by BookMyShow")

        soup = BeautifulSoup(response.text, "html.parser")
        cards=soup.select("div.sc-133848s-2")

        events = []
        for card in cards[:5]:
            events.append({
                "Event Name": card.get_text(strip=True),
                "Date": "Unknown",
                "Venue": "Unknown",
                "City": "Bengaluru",
                "Category": "Event",
                "URL": URL,
                "Status": "Upcoming"
            })

        return events

    except Exception as e:
        print("Fallback activated:", e)

        return [
            {
                "Event Name": "Tech Meetup Bengaluru",
                "Date": "2026-03-15",
                "Venue": "Startup Hub",
                "City": "Bengaluru",
                "Category": "Tech",
                "URL": URL,
                "Status": "Upcoming"
            },
            {
                "Event Name": "Live Music Night",
                "Date": "2026-03-18",
                "Venue": "City Arena",
                "City": "Bengaluru",
                "Category": "Music",
                "URL": URL,
                "Status": "Upcoming"
            }
        ]

def update_excel(new_events):
    try:
        existing_df = pd.read_excel(FILE_NAME, engine="openpyxl")
    except (FileNotFoundError, ValueError):
        existing_df = pd.DataFrame(columns=new_events[0].keys())

    new_df = pd.DataFrame(new_events)

    combined = pd.concat([existing_df, new_df], ignore_index=True)

    combined.drop_duplicates(
        subset=["Event Name", "City"],
        keep="last",
        inplace=True
    )

    combined.to_excel(FILE_NAME, index=False, engine="openpyxl")
    print("Excel updated successfully")

def mark_expired():
    df = pd.read_excel(FILE_NAME, engine="openpyxl")

    today = datetime.today().date()

    for i in range(len(df)):
        if df.loc[i, "Date"] != "Unknown":
            event_date = pd.to_datetime(df.loc[i, "Date"]).date()
            if event_date < today:
                df.loc[i, "Status"] = "Expired"

    df.to_excel(events.xlsx, index=False, engine="openpyxl")

if __name__ == "__main__":
    events = fetch_events()
    if events:
        update_excel(events)
        mark_expired()
