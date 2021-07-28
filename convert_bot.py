#!/usr/bin/env python

"""
Webex bot
Send a PPTX file to the bot and the bot will then convert all theme colors to RGB colors and send back
the converted deck
"""

import cgi
import hashlib
import hmac
import logging
import os
import random
import tempfile
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from typing import List, Callable

import requests
import webexteamssdk
from dotenv import load_dotenv
from flask import Flask, request
from webexteamssdk import Message
from webexteamssdk import WebexTeamsAPI

import ngrokhelper
from webex_convert import convert_pptx_to_rgb

load_dotenv()

log = logging.getLogger(__name__)


class BotMessageProcessor:

    def __init__(self, access_token: str, **kwargs):
        self.access_token = access_token

    @contextmanager
    def get_file(self, file_url: str, room_id: str, api: WebexTeamsAPI) -> requests.Response:
        with requests.Session() as session:
            while True:
                with session.get(url=file_url, headers={'Authorization': f'Bearer {api.access_token}'},
                                 stream=True) as response:
                    if response.status_code == 423:
                        retry = int(response.headers.get('retry-after', '10'))
                        cd_header = response.headers.get('content-disposition', None)
                        _, params = cgi.parse_header(cd_header)
                        file_name = params['filename']
                        log.debug(f'Waiting {retry} seconds for {file_name} to become available')
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
                        log.debug(f'download failed: {response.status_code}/{response.reason}')
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
                    log.debug(f'downloading {full_path}')
                    with open(full_path, mode='wb') as file:
                        for chunk in response.iter_content(chunk_size=2 * 1024 * 1024):
                            log.debug(f'{file_name}: got chunk, {len(chunk)} bytes')
                            file.write(chunk)
                rgb_path = f'{os.path.splitext(full_path)[0]}_rgb.pptx'
                api.messages.create(roomId=message.roomId,
                                    text=f'Converting {file_name}')
                convert_pptx_to_rgb(full_path, rgb_path)
                api.messages.create(roomId=message.roomId,
                                    text='Here is the converted PPTX',
                                    files=[rgb_path])


MessageCallback = Callable[[webexteamssdk.Message], None]


class BotWebhook(Flask):
    def __init__(self, access_token: str,
                 message_callback: MessageCallback,
                 allowed_emails: List[str] = None):
        super().__init__(import_name=__name__)
        self.access_token = access_token
        self._message_callback = message_callback
        self._allowed_emails = set(allowed_emails or list())
        self._pool = ThreadPoolExecutor()

        heroku_name = os.getenv('HEROKU_NAME')
        if heroku_name is None:
            log.debug('not running on Heroku. Creating Ngrok tunnel')
            ngrok = ngrokhelper.NgrokHelper(port=5000)
            bot_url = ngrok.start()
        else:
            log.debug(f'running on heroku as {heroku_name}')
            bot_url = f'https://{heroku_name}.herokuapp.com'
        log.debug(f'Webhook URL: {bot_url}')
        self.add_url_rule(
            "/", "index", self.process_incoming_message, methods=["POST"]
        )
        self._secret = str(uuid.uuid4())
        self._api = WebexTeamsAPI(access_token=self.access_token)
        me = self._api.people.me()
        self.me_id = me.id
        self._pool.submit(self.setup_hooks, url=bot_url)

    def setup_hooks(self, url: str):
        api = WebexTeamsAPI(access_token=self.access_token)
        while True:
            hooks: List[webexteamssdk.Webhook] = list(api.webhooks.list())
            log.debug(f'found {len(hooks)} webhooks')
            if not hooks:
                # create one
                log.debug(f'create new webhook')
                api.webhooks.create(name='messages.created',
                                    targetUrl=url,
                                    resource='messages',
                                    event='created',
                                    secret=self._secret)
            else:
                for hook in hooks[1:]:
                    log.debug(f'trying to delete webhook {hook.id}')
                    try:
                        api.webhooks.delete(webhookId=hook.id)
                    except webexteamssdk.ApiError:
                        pass
                api.webhooks.update(webhookId=hooks[0].id,
                                    name='messages.created',
                                    targetUrl=url)
                # set secret
                self._secret = hooks[0].secret
                if len(hooks) == 1:
                    break
            s = random.randint(1, 5)
            log.debug(f'Waiting {s} s before validating hooks')
            time.sleep(s)

        log.debug('Done setting up the web hook')

    def process_incoming_message(self):
        """
        Handle message.created events
        :return:
        """
        # validate signature
        raw = request.get_data()
        # Let's create the SHA1 signature
        # based on the request body JSON (raw) and our passphrase (key)
        hashed = hmac.new(self._secret.encode(), raw, hashlib.sha1)
        validatedSignature = hashed.hexdigest()
        signature = request.headers.get('X-Spark-Signature')
        if signature != validatedSignature:
            log.warning('signature mismatch: ignore')
            return 'ok'
        data = request.json
        event = webexteamssdk.WebhookEvent(data)
        if event.data.personId == self.me_id:
            log.debug(f'ignoring message from self')
            return 'ok'
        email = event.data.personEmail
        if self._allowed_emails and email not in self._allowed_emails:
            log.debug(f'{email} not allowed. Skipping message')
        # get message
        message = self._api.messages.get(messageId=event.data.id)
        self._pool.submit(self._message_callback, message=message)
        return 'ok'

    def run(self, host: str = '0.0.0.0', port: int = 5000):
        super().run(host=host, port=port)


# base class for bot communication
# bot_base = BotSocket
bot_base = BotWebhook


class PPTBot(bot_base, BotMessageProcessor):

    def __init__(self):
        access_token = os.getenv('BOT_ACCESS_TOKEN')
        if access_token is None:
            raise Exception('access token needs to be defined in env variable BOT_ACCESS_TOKEN')

        self.access_token = access_token
        super().__init__(access_token=access_token, message_callback=self.process_message_sync)


logging.basicConfig(level=logging.DEBUG, format='%(asctime)s [%(process)d] %(threadName)s %(levelname)s %(module)s %('
                                                'message)s')
bot = PPTBot()

if __name__ == '__main__':
    bot.run()
