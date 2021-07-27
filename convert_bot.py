#!/usr/bin/env python

"""
Webex bot
Send a PPTX file to the bot and the bot will then convert all theme colors to RGB colors and send back
the converted deck
"""

from botsocket import BotSocket
from dotenv import load_dotenv
from webexteamssdk import Message
from webexteamssdk import WebexTeamsAPI
import os
import logging
import asyncio
import tempfile
import requests
import cgi
import time
from webex_convert import convert_pptx_to_rgb
from contextlib import contextmanager

load_dotenv()

class PPTBot(BotSocket):
    def __init__(self):
        access_token = os.getenv('BOT_ACCESS_TOKEN')
        self.access_token = access_token
        if access_token is None:
            raise Exception('access token needs to be defined in env variable BOT_ACCESS_TOKEN')
        super().__init__(access_token=access_token, message_callback=self.process_message)

    async def process_message(self, message: Message):
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(None, self.process_message_sync, message)

    @contextmanager
    def get_file(self, file_url: str, room_id: str, api:WebexTeamsAPI)->requests.Response:
        with requests.Session() as session:
            while True:
                with session.get(url=file_url, headers={'Authorization': f'Bearer {self.access_token}'},
                                 stream=True) as response:
                    if response.status_code == 423:
                        retry = int(response.headers.get('retry-after', '10'))
                        cd_header = response.headers.get('content-disposition', None)
                        _, params = cgi.parse_header(cd_header)
                        file_name = params['filename']
                        logging.debug(f'Waiting {retry} seconds for {file_name} to become available')
                        api.messages.create(roomId=room_id,
                                            text=f'Waiting {retry} seconds for {file_name} to become available')
                        time.sleep(retry)
                        continue
                    yield response
                    break

    def process_message_sync(self, message: Message):
        api = WebexTeamsAPI(access_token=self.access_token)
        # we need a message w/ attachment
        if not message.files:
            api.messages.create(roomId=message.roomId, text='Send me a PPTX file and I will return a converted version')
            return
        for file_url in message.files:
            with tempfile.TemporaryDirectory() as tempdir:
                with self.get_file(file_url=file_url, room_id=message.roomId, api=api) as response:
                    if response.status_code != 200:
                        logging.debug(f'download failed: {response.status_code}/{response.reason}')
                        continue
                    cd_header = response.headers.get('content-disposition', None)
                    _, params = cgi.parse_header(cd_header)
                    file_name = params['filename']
                    _, ext = os.path.splitext(file_name)
                    if ext.lower() != '.pptx':
                        api.messages.create(roomId=message.roomId,
                                            text=f'Send me a PPTX (and not {ext.upper()[1:]} file and I will '
                                                 f'return a converted version')
                        continue
                    full_path = os.path.join(tempdir, file_name)
                    api.messages.create(roomId=message.roomId,
                                        text=f'Downloading {file_name}')
                    logging.debug(f'downloading {full_path}')
                    with open(full_path, mode='wb') as file:
                        for chunk in response.iter_content(chunk_size=2*1024*1024):
                            logging.debug(f'{file_name}: got chunk, {len(chunk)} bytes')
                            file.write(chunk)
                rgb_path = f'{os.path.splitext(full_path)[0]}_rgb.pptx'
                api.messages.create(roomId=message.roomId,
                                    text=f'Converting {file_name}')
                convert_pptx_to_rgb(full_path, rgb_path)
                api.messages.create(roomId=message.roomId,
                                    text='Here is the converted PPTX',
                                    files=[rgb_path])


if __name__ == '__main__':
    print('Starting')
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(threadName)s %(levelname)s %(module)s %(message)s')
    bot = PPTBot()
    bot.run()
