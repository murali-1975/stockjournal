import telethon.tl.functions as funcs
import telethon.tl.types as types

def find():
    print("--- Searching for 'Topic' in telethon.tl.functions ---")
    for m in dir(funcs):
        if "Topic" in m:
            print(f"funcs.{m}")
            
    print("\n--- Searching for 'Topic' in telethon.tl.functions.channels ---")
    import telethon.tl.functions.channels as channels
    for m in dir(channels):
        if "Topic" in m:
            print(f"channels.{m}")

if __name__ == "__main__":
    find()
