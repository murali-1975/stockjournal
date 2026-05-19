from telethon.sync import TelegramClient
from telethon.sessions import StringSession

API_ID = 35895979
API_HASH = '59e15f88db969d701d7a03d814e94664'
SESSION_STR = '1BVtsOIgBu78XP2swCqevYQJgpiPlG5glTPnSWNU3eJ_skaW4abp10ZnDC6AaFDyxXunMr39fnDu7teo6F2UtpL69kV6pXJCAka9VCsF6_Q-qlc9yuyaE91xvoRBO3tRMRC4hTC2sd1mcwoMeE0al7mba5ih2bmDolDRivtob0xBWJYi85bElCFzEIvQZ58IYm8AH8kwCuAmYPa4-lBoKa9s8pPBqNJ08g63cbf5xVNv2agXw0csP0QhR9xVOpGxUHfmuEQ9V4GbedNpU8JsxzGs9j0Jz5KGqZvVynR5tje4ad-KQtJWkoDARHMRdkwYKTlVKIWkeSK_GbFr4BLQHaOLRCUmTaxE='

def list_dialogs():
    with TelegramClient(StringSession(SESSION_STR), API_ID, API_HASH) as client:
        print("\n--- Listing All Dialogs ---")
        for dialog in client.iter_dialogs():
            print(f"- {dialog.name} (ID: {dialog.id}) Type: {'Group' if dialog.is_group else 'Channel' if dialog.is_channel else 'User'}")
            
if __name__ == "__main__":
    list_dialogs()
