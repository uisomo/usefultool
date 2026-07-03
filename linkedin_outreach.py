"""
LinkedIn outreach helper
========================
For each target profile:
  1. Opens the profile and checks their current company.
     -> SKIPS anyone at MUFG (or Mitsubishi UFJ / 三菱UFJ etc.)
  2. Opens their recent activity and grabs the text of their latest post.
  3. Uses AI (OpenRouter) to draft:
       - a short comment for that post
       - a casual connection request note / message that references the post
  4. If not connected  -> sends a connection request with the note.
     If already 1st    -> sends a direct message instead.
     Optionally (--comment) also posts the comment on their latest post.

Setup
-----
  pip install selenium openai python-dotenv
  .env file with:  OPENROUTER_API_KEY=sk-or-...
  targets file (one LinkedIn profile URL per line), e.g. linkedin_targets.txt

Usage
-----
  python linkedin_outreach.py linkedin_targets.txt            # dry run: only prints drafts
  python linkedin_outreach.py linkedin_targets.txt --send     # asks y/n before each send
  python linkedin_outreach.py linkedin_targets.txt --send --comment
  python linkedin_outreach.py linkedin_targets.txt --send --yes   # no confirmation (careful!)

NOTES / WARNINGS
----------------
* Automating LinkedIn is against LinkedIn's Terms of Service. Accounts that
  send too many automated invites/messages get restricted or banned.
  Keep MAX_PER_RUN low (default 15) and keep the random delays.
* LinkedIn changes its HTML often. If a selector stops working, fix it in
  the SELECTORS section below.
* You log in manually in the opened browser window (safer than storing
  your password in a script).
"""

import argparse
import json
import os
import random
import re
import sys
import time

from dotenv import load_dotenv
from openai import OpenAI
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# ----------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------
MAX_PER_RUN = 15                 # hard cap on sends per run - keep this low
DELAY_RANGE = (45, 100)          # random seconds to wait between profiles
NOTE_CHAR_LIMIT = 280            # LinkedIn invite notes max out at 300 chars
AI_MODEL = "deepseek/deepseek-r1-distill-llama-70b"

# Anyone whose profile top-card mentions these is skipped
COMPANY_BLOCKLIST = [
    "mufg",
    "mitsubishi ufj",
    "三菱UFJ",
    "三菱ＵＦＪ",
]

# ----------------------------------------------------------------------
# AI draft generation (OpenRouter, same setup as the `openrouter` note)
# ----------------------------------------------------------------------
load_dotenv()
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.getenv("OPENROUTER_API_KEY"),
)

PROMPT = """You help write short, casual, friendly LinkedIn outreach in the
same language the post is written in (Japanese post -> Japanese reply).

The person's name: {name}
Their headline: {headline}
Their latest LinkedIn post:
---
{post}
---

Write JSON with exactly these keys:
  "comment": a 1-2 sentence comment to leave on the post. Specific to the
             post content, warm, no flattery overload, no hashtags.
  "message": a casual connection request note (max {limit} characters).
             Mention something concrete from their post, then casually ask
             to connect. No hard selling, no "Dear", keep it human.

Return ONLY the JSON object, nothing else.
"""


def draft_with_ai(name, headline, post):
    completion = client.chat.completions.create(
        extra_headers={
            "HTTP-Referer": os.getenv("SITE_URL", "http://localhost"),
            "X-Title": os.getenv("SITE_NAME", "LinkedIn Outreach"),
        },
        model=AI_MODEL,
        messages=[{
            "role": "user",
            "content": PROMPT.format(
                name=name, headline=headline, post=post[:2000], limit=NOTE_CHAR_LIMIT
            ),
        }],
    )
    raw = completion.choices[0].message.content
    match = re.search(r"\{.*\}", raw, re.DOTALL)  # tolerate text around the JSON
    data = json.loads(match.group(0)) if match else {"comment": "", "message": raw}
    data["message"] = data.get("message", "")[:NOTE_CHAR_LIMIT]
    return data


# ----------------------------------------------------------------------
# SELECTORS - the parts most likely to need fixing when LinkedIn changes
# ----------------------------------------------------------------------
SEL_NAME = "h1"
SEL_HEADLINE = "div.text-body-medium.break-words"
SEL_TOPCARD = "section.artdeco-card"            # first card = profile top card
SEL_POST = "div.feed-shared-update-v2"          # first one on activity page
SEL_POST_TEXT = "div.update-components-text"
SEL_INVITE_NOTE = "textarea#custom-message"
SEL_MSG_BOX = "div.msg-form__contenteditable"
SEL_MSG_SEND = "button.msg-form__send-button"
SEL_COMMENT_BOX = "div.comments-comment-box__form div[contenteditable='true']"


def find(driver, css, many=False):
    els = driver.find_elements(By.CSS_SELECTOR, css)
    if many:
        return els
    return els[0] if els else None


def click_button_by_text(driver, texts):
    """Click the first visible button whose aria-label or text contains any of `texts`."""
    for btn in driver.find_elements(By.TAG_NAME, "button"):
        label = ((btn.get_attribute("aria-label") or "") + " " + (btn.text or "")).lower()
        if btn.is_displayed() and any(t.lower() in label for t in texts):
            driver.execute_script("arguments[0].click();", btn)
            return True
    return False


# ----------------------------------------------------------------------
# Profile handling
# ----------------------------------------------------------------------
def read_profile(driver, url):
    driver.get(url)
    time.sleep(random.uniform(4, 7))
    name_el = find(driver, SEL_NAME)
    headline_el = find(driver, SEL_HEADLINE)
    topcard = find(driver, SEL_TOPCARD)
    return {
        "url": url.rstrip("/"),
        "name": name_el.text.strip() if name_el else "",
        "headline": headline_el.text.strip() if headline_el else "",
        "topcard_text": topcard.text if topcard else driver.page_source[:5000],
        "is_first_degree": "· 1st" in (topcard.text if topcard else "")
        or "· １度" in (topcard.text if topcard else ""),
    }


def works_at_blocked_company(profile):
    haystack = (profile["headline"] + " " + profile["topcard_text"]).lower()
    return any(b.lower() in haystack for b in COMPANY_BLOCKLIST)


def get_latest_post(driver, profile_url):
    driver.get(profile_url + "/recent-activity/all/")
    time.sleep(random.uniform(4, 7))
    post = find(driver, SEL_POST)
    if not post:
        return None
    text_el = post.find_elements(By.CSS_SELECTOR, SEL_POST_TEXT)
    return (text_el[0].text if text_el else post.text).strip() or None


# ----------------------------------------------------------------------
# Actions
# ----------------------------------------------------------------------
def send_connection_request(driver, profile, note):
    driver.get(profile["url"])
    time.sleep(random.uniform(3, 5))
    # Connect is sometimes hidden under "More"
    if not click_button_by_text(driver, ["Invite", "Connect", "つながり"]):
        click_button_by_text(driver, ["More actions", "More", "その他"])
        time.sleep(1.5)
        if not click_button_by_text(driver, ["Invite", "Connect", "つながり"]):
            print("  !! Connect button not found - maybe pending already. Skipping.")
            return False
    time.sleep(2)
    if click_button_by_text(driver, ["Add a note", "メモを追加"]):
        time.sleep(1.5)
        box = find(driver, SEL_INVITE_NOTE)
        if box:
            box.send_keys(note)
            time.sleep(1)
    return click_button_by_text(driver, ["Send invitation", "Send", "送信"])


def send_direct_message(driver, profile, message):
    driver.get(profile["url"])
    time.sleep(random.uniform(3, 5))
    if not click_button_by_text(driver, ["Message", "メッセージ"]):
        print("  !! Message button not found. Skipping.")
        return False
    time.sleep(2.5)
    box = find(driver, SEL_MSG_BOX)
    if not box:
        print("  !! Message box not found. Skipping.")
        return False
    box.click()
    box.send_keys(message)
    time.sleep(1)
    return click_button_by_text(driver, ["Send", "送信"]) or (
        box.send_keys(Keys.CONTROL, Keys.RETURN) or True
    )


def post_comment_on_latest(driver, profile, comment):
    driver.get(profile["url"] + "/recent-activity/all/")
    time.sleep(random.uniform(4, 6))
    post = find(driver, SEL_POST)
    if not post:
        return False
    # open the comment box on the first post
    for btn in post.find_elements(By.TAG_NAME, "button"):
        label = (btn.get_attribute("aria-label") or "").lower()
        if "comment" in label or "コメント" in label:
            driver.execute_script("arguments[0].click();", btn)
            break
    time.sleep(2)
    box = find(driver, SEL_COMMENT_BOX)
    if not box:
        print("  !! Comment box not found. Skipping comment.")
        return False
    box.click()
    box.send_keys(comment)
    time.sleep(1)
    return click_button_by_text(driver, ["Post comment", "Comment", "投稿", "コメント"])


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="LinkedIn outreach helper")
    ap.add_argument("targets", help="text file with one LinkedIn profile URL per line")
    ap.add_argument("--send", action="store_true", help="actually send (default: dry run)")
    ap.add_argument("--comment", action="store_true", help="also comment on their latest post")
    ap.add_argument("--yes", action="store_true", help="don't ask y/n before each send")
    args = ap.parse_args()

    with open(args.targets, encoding="utf-8") as f:
        urls = [u.strip() for u in f if u.strip().startswith("http")]
    if not urls:
        sys.exit("No profile URLs found in " + args.targets)

    driver = webdriver.Chrome()
    driver.get("https://www.linkedin.com/login")
    input("\n>> Log in to LinkedIn in the browser window, then press Enter here... ")

    sent = 0
    for i, url in enumerate(urls, 1):
        if sent >= MAX_PER_RUN:
            print(f"\nReached MAX_PER_RUN ({MAX_PER_RUN}). Stopping - run again tomorrow.")
            break

        print(f"\n[{i}/{len(urls)}] {url}")
        try:
            profile = read_profile(driver, url)
            print(f"  {profile['name']} | {profile['headline'][:60]}")

            if works_at_blocked_company(profile):
                print("  -> works at MUFG(-related company). SKIP.")
                continue

            post = get_latest_post(driver, profile["url"])
            if not post:
                print("  -> no recent post found. SKIP.")
                continue
            print(f"  latest post: {post[:100]}...")

            draft = draft_with_ai(profile["name"], profile["headline"], post)
            print(f"  --- comment draft ---\n  {draft['comment']}")
            print(f"  --- message draft ---\n  {draft['message']}")

            if not args.send:
                print("  (dry run - nothing sent)")
                continue

            if not args.yes:
                if input("  Send this? [y/N] ").strip().lower() != "y":
                    print("  skipped by you.")
                    continue

            if args.comment:
                ok = post_comment_on_latest(driver, profile, draft["comment"])
                print("  comment posted." if ok else "  comment failed.")

            if profile["is_first_degree"]:
                ok = send_direct_message(driver, profile, draft["message"])
                print("  message sent." if ok else "  message failed.")
            else:
                ok = send_connection_request(driver, profile, draft["message"])
                print("  connection request sent." if ok else "  connection request failed.")

            if ok:
                sent += 1

        except Exception as e:
            print(f"  !! error on this profile: {e}")

        wait = random.uniform(*DELAY_RANGE)
        print(f"  waiting {wait:.0f}s (human pace)...")
        time.sleep(wait)

    print(f"\nDone. Sent {sent} of {len(urls)} targets.")
    driver.quit()


if __name__ == "__main__":
    main()
