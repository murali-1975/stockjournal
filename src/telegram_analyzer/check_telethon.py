import telethon.tl.functions.channels as channels

def check():
    print("--- Functions in telethon.tl.functions.channels ---")
    for name in dir(channels):
        if "Topic" in name:
            print(name)

if __name__ == "__main__":
    check()
