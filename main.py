from telethon import TelegramClient, events, utils
import openpyxl
import os
from dotenv import load_dotenv

load_dotenv()

api_id = os.getenv('API_ID')
api_hash = os.getenv('API_HASH')
limit = int(os.getenv('LIMIT'))

print(limit)
excel_file = 'telegram_data.xlsx'

media_folder = 'media'

if not os.path.exists(media_folder):
    os.makedirs(media_folder)


async def main():
    client = TelegramClient('my_session', api_id, api_hash)
    await client.start()

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['channel_name', 'post_text', 'post_media'])

    try:
        with open('channels.txt', 'r') as f:
            for channel_link in f:
                print(f"Обработка канала: {channel_link}")
                channel_link = channel_link.strip() # Remove leading/trailing whitespace
                try:
                     channel_entity = await client.get_entity(channel_link)
                     async for message in client.iter_messages(channel_entity, limit=limit):
                         
                         post_text = message.text if message.text else ""
                         post_media = ""
                         
                         if message.photo:
                            file_path = await message.download_media(file=media_folder)
                            post_media = os.path.join(media_folder, os.path.basename(file_path))
                         elif message.media:
                            file_path = await message.download_media(file=media_folder)
                            post_media = os.path.join(media_folder, os.path.basename(file_path))
                         
                         sheet.append([channel_entity.title, post_text, post_media])
                     print(f'Канал {channel_link} обработан')
                except Exception as e:
                    print(f"Ошибка при обработке канала {channel_link}: {e}")


    except FileNotFoundError:
        print("channels.txt not found.")
    finally:
        workbook.save(excel_file)
        await client.disconnect()

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
