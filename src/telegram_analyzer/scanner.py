import asyncio
import argparse
import pandas as pd
import gspread
from datetime import datetime, timedelta, timezone
from telethon import TelegramClient
from telethon.sessions import StringSession
from telethon.tl.functions.messages import GetForumTopicsRequest
from google.oauth2.service_account import Credentials
import re

# --- CONFIGURATION ---
API_ID = 35895979
API_HASH = '59e15f88db969d701d7a03d814e94664'
SESSION_STR = '1BVtsOIgBu78XP2swCqevYQJgpiPlG5glTPnSWNU3eJ_skaW4abp10ZnDC6AaFDyxXunMr39fnDu7teo6F2UtpL69kV6pXJCAka9VCsF6_Q-qlc9yuyaE91xvoRBO3tRMRC4hTC2sd1mcwoMeE0al7mba5ih2bmDolDRivtob0xBWJYi85bElCFzEIvQZ58IYm8AH8kwCuAmYPa4-lBoKa9s8pPBqNJ08g63cbf5xVNv2agXw0csP0QhR9xVOpGxUHfmuEQ9V4GbedNpU8JsxzGs9j0Jz5KGqZvVynR5tje4ad-KQtJWkoDARHMRdkwYKTlVKIWkeSK_GbFr4BLQHaOLRCUmTaxE='

PARENT_GROUP_ID = -1001754401822 
TECHNOFUNDA_CHANNEL_ID = -1001300781517 

# Normalized names for matching
TARGET_TOPICS_NORMALIZED = [
    "satellite portfolio discussion",
    "success stories",
    "us investing",
    "rising star incubation",
    "message from technofunda",
    "core portfolio discussion",
    "interesting charts",
    "ipo discussion",
    "learn from mistakes"
]

CORE_WATCHLIST_ID = "1aQ0xEBZnuAFuECMZZwVUlqh4u8-TNFU1q5eIvIYdE3M"
SAT_WATCHLIST_ID = "1aQD5peZMqoo3odx47d-lafvP-CrmxmrsxiJSaFGW_rc"
OUTPUT_SHEET_ID = "1BBE2GjZe3dC_SaM_bSygZgtLLKMaDHVj2loHSNoR5Wo"

def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    return gspread.authorize(creds)

def load_watchlists(client):
    watchlists = {} 
    try:
        sheets = [(CORE_WATCHLIST_ID, "Core"), (SAT_WATCHLIST_ID, "Satellite")]
        for sheet_id, cat in sheets:
            sh = client.open_by_key(sheet_id)
            records = sh.get_worksheet(0).get_all_records()
            for r in records:
                sym = str(r.get('Stock Symbol', r.get('Stock', ''))).strip()
                comp = str(r.get('Company Name', '')).strip()
                if not sym or sym == 'Stock Symbol': continue
                patterns = [sym]
                if comp:
                    patterns.append(comp)
                    first_word = comp.split(' ')[0]
                    if len(first_word) >= 5 and first_word.upper() not in ['INDIA', 'LIMITED', 'CORP', 'POWER']:
                        patterns.append(first_word)
                watchlists[sym] = {"cat": cat, "patterns": list(set(patterns))}
    except Exception as e:
        print(f"Error loading Watchlists: {e}")
    return watchlists

def create_regex_pattern(text):
    escaped = re.escape(text)
    flexible = escaped.replace(r'\-', r'[\s\-]*').replace(r'\ ', r'[\s\-]*')
    return rf'\b{flexible}\b'

async def run_scanner(dump_all=False, days=7):
    gs_client = get_gspread_client()
    watchlists = load_watchlists(gs_client)
    
    try:
        output_ws = gs_client.open_by_key(OUTPUT_SHEET_ID).get_worksheet(0)
    except Exception as e:
        print(f"Error opening output sheet: {e}")
        return
    
    print(f"Connecting to Telegram (Deep Topic Scanning: ON)...")
    async with TelegramClient(StringSession(SESSION_STR), API_ID, API_HASH) as client:
        group_entity = await client.get_entity(PARENT_GROUP_ID)
        result = await client(GetForumTopicsRequest(peer=group_entity, offset_date=None, offset_id=0, offset_topic=0, limit=100))
        
        # Determine which topics to scan
        topics_to_scan = []
        for topic in result.topics:
            if dump_all or topic.title.lower() in TARGET_TOPICS_NORMALIZED:
                topics_to_scan.append(topic)

        channel_entity = await client.get_entity(TECHNOFUNDA_CHANNEL_ID)
        time_limit = datetime.now(timezone.utc) - timedelta(days=days)
        new_signals = []
        
        compiled_patterns = {}
        for sym, data in watchlists.items():
            compiled_patterns[sym] = [re.compile(create_regex_pattern(p), re.IGNORECASE) for p in data['patterns']]

        # 1. Scan Topics Individually (Much more robust)
        for topic in topics_to_scan:
            print(f"Scanning Topic: {topic.title}...")
            topic_msgs_cache = {}
            topic_msgs_list = []
            
            async for message in client.iter_messages(group_entity, limit=1000, reply_to=topic.id):
                if message.date < time_limit: break
                topic_msgs_cache[message.id] = message
                topic_msgs_list.append(message)
            
            for message in topic_msgs_list:
                content = message.text if message.text else ""
                if not content: continue
                
                # Threading Context
                parent_id = getattr(message.reply_to, 'reply_to_msg_id', None) if message.reply_to else None
                combined_content = content
                if parent_id and parent_id in topic_msgs_cache:
                    parent_msg = topic_msgs_cache[parent_id]
                    parent_text = parent_msg.text if parent_msg.text else "[Media/Image]"
                    combined_content = f"[ORIGINAL]: {parent_text}\n--- REPLY ---\n{content}"

                matched_sym = ""
                matched_cat = ""
                for sym, data in watchlists.items():
                    if any(p.search(content) for p in compiled_patterns[sym]):
                        matched_sym = sym
                        matched_cat = data['cat']
                        break
                
                if dump_all or matched_sym:
                    new_signals.append([message.date.strftime('%Y-%m-%d %H:%M'), matched_sym, matched_cat, combined_content, topic.title])

        # 2. Scan Independent Channel
        print(f"Scanning Channel: {channel_entity.title}...")
        channel_msgs_cache = {}
        channel_msgs_list = []
        async for message in client.iter_messages(channel_entity, limit=500):
            if message.date < time_limit: break
            channel_msgs_cache[message.id] = message
            channel_msgs_list.append(message)

        for message in channel_msgs_list:
            content = message.text if message.text else ""
            if not content: continue
            
            parent_id = getattr(message.reply_to, 'reply_to_msg_id', None) if message.reply_to else None
            combined_content = content
            if parent_id and parent_id in channel_msgs_cache:
                parent_msg = channel_msgs_cache[parent_id]
                parent_text = parent_msg.text if parent_msg.text else "[Media/Image]"
                combined_content = f"[ORIGINAL]: {parent_text}\n--- REPLY ---\n{content}"

            matched_sym = ""
            matched_cat = ""
            for sym, data in watchlists.items():
                if any(p.search(content) for p in compiled_patterns[sym]):
                    matched_sym = sym
                    matched_cat = data['cat']
                    break
            
            if dump_all or matched_sym:
                new_signals.append([message.date.strftime('%Y-%m-%d %H:%M'), matched_sym, matched_cat, combined_content, channel_entity.title])

        if new_signals:
            new_signals.sort(key=lambda x: x[0])
            print(f"Uploading {len(new_signals)} row(s)...")
            batch_size = 500
            for i in range(0, len(new_signals), batch_size):
                output_ws.append_rows(new_signals[i:i+batch_size])
            print("Done!")
        else:
            print("No data found.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Stock Journal Telegram Scanner")
    parser.add_argument("--dump-all", action="store_true", help="Dump all messages irrespective of watchlist")
    parser.add_argument("--days", type=int, default=7, help="Number of days to look back")
    args = parser.parse_args()
    
    asyncio.run(run_scanner(dump_all=args.dump_all, days=args.days))
