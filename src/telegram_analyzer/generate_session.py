from telethon.sync import TelegramClient
from telethon.sessions import StringSession

# Your Telegram API Credentials
API_ID = 35895979
API_HASH = '59e15f88db969d701d7a03d814e94664'

print("\n--- Telegram Session Generator ---")
print("This script will help you log in once and generate a permanent session string.")
print("You will need to enter your phone number and the OTP sent to your Telegram app.\n")

try:
    with TelegramClient(StringSession(), API_ID, API_HASH) as client:
        session_str = client.session.save()
        print("\n" + "="*50)
        print("SUCCESS! HERE IS YOUR SESSION STRING:")
        print("="*50)
        print(session_str)
        print("="*50)
        print("\nIMPORTANT: Copy the entire string above and keep it safe.")
        print("We will use this string in your Google Cloud Run environment variables.")
        print("Do NOT share this string with anyone, as it gives access to your Telegram account.\n")
except Exception as e:
    print(f"\nError: {e}")
