import asyncio
from telethon import TelegramClient
from telethon.sessions import StringSession
from telethon.tl.functions.messages import GetForumTopicsRequest

API_ID = 35895979
API_HASH = '59e15f88db969d701d7a03d814e94664'
SESSION_STR = '1BVtsOIgBu78XP2swCqevYQJgpiPlG5glTPnSWNU3eJ_skaW4abp10ZnDC6AaFDyxXunMr39fnDu7teo6F2UtpL69kV6pXJCAka9VCsF6_Q-qlc9yuyaE91xvoRBO3tRMRC4hTC2sd1mcwoMeE0al7mba5ih2bmDolDRivtob0xBWJYi85bElCFzEIvQZ58IYm8AH8kwCuAmYPa4-lBoKa9s8pPBqNJ08g63cbf5xVNv2agXw0csP0QhR9xVOpGxUHfmuEQ9V4GbedNpU8JsxzGs9j0Jz5KGqZvVynR5tje4ad-KQtJWkoDARHMRdkwYKTlVKIWkeSK_GbFr4BLQHaOLRCUmTaxE='
GROUP_ID = -1001754401822

async def debug_topics():
    async with TelegramClient(StringSession(SESSION_STR), API_ID, API_HASH) as client:
        entity = await client.get_entity(GROUP_ID)
        print("\n--- All Topics Found in Group ---")
        result = await client(GetForumTopicsRequest(
            peer=entity,
            offset_date=None,
            offset_id=0,
            offset_topic=0,
            limit=100
        ))
        for t in result.topics:
            print(f"- {t.title} (ID: {t.id})")

if __name__ == "__main__":
    asyncio.run(debug_topics())
