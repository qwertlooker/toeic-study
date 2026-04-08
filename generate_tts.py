import asyncio
import edge_tts
import os

TEXT_DIR = r'e:\English\02_Listening_P3P4\articles_text'
OUTPUT_DIR = r'e:\English\02_Listening_P3P4\audio'
os.makedirs(OUTPUT_DIR, exist_ok=True)

VOICE = "en-US-JennyNeural"
RATE = "-10%"

async def generate_tts(text_file, output_file):
    communicate = edge_tts.Communicate(text=open(text_file, 'r', encoding='utf-8').read(), voice=VOICE, rate=RATE)
    await communicate.save(output_file)
    print(f"Generated: {os.path.basename(output_file)}")

async def main():
    tasks = []
    for i in range(1, 21):
        text_file = os.path.join(TEXT_DIR, f"article_{i:02d}.txt")
        output_file = os.path.join(OUTPUT_DIR, f"article_{i:02d}.mp3")
        if os.path.exists(text_file):
            tasks.append(generate_tts(text_file, output_file))
    await asyncio.gather(*tasks)
    print(f"\nAll {len(tasks)} audio files generated!")

asyncio.run(main())
