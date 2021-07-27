import asyncio
import base64
import functools
import json
import logging
import os
from typing import Optional, Callable, List, Coroutine

import aiohttp
import webexteamssdk

WDM_DEVICES = 'https://wdm-a.wbx2.com/wdm/api/v1/devices'

log = logging.getLogger(__name__)

MessageCallback = Callable[[webexteamssdk.Message], None]


class BotSocket:
    """
    Bot helper based on Webex Teams device registration and Websocket
    """

    def __init__(self,
                 access_token: str,
                 message_callback: Optional[MessageCallback],
                 allowed_emails: List[str] = None) -> None:
        self._token = access_token
        self._device_name = os.path.basename(os.path.splitext(__file__)[0])
        self._message_callback = message_callback
        self._session: Optional[aiohttp.ClientSession] = None
        self._api = webexteamssdk.WebexTeamsAPI(access_token=access_token)
        self._allowed_emails = set(allowed_emails or list())

    @property
    def auth(self):
        return f'Bearer {self._token}'

    async def request(self, method: str, url: str, **kwargs):
        headers = kwargs.get('headers', dict())
        headers['Authorization'] = self.auth
        kwargs['headers'] = headers
        async with self._session.request(method=method, url=url, **kwargs) as r:
            r.raise_for_status()
            result = await r.json()
        return result

    async def get(self, url, **kwargs):
        return await self.request(method='GET', url=url, **kwargs)

    async def post(self, url, **kwargs):
        return await self.request(method='POST', url=url, **kwargs)

    async def find_device(self) -> Optional[dict]:
        try:
            r = await self.get(url=WDM_DEVICES)
            devices = r['devices']
            if len(devices) > 1:
                log.warning(f'Found {len(devices)} devices: {", ".join(d["name"] for d in devices)}')
            device = next((d for d in devices if d['name'] == self._device_name), None)
        except aiohttp.ClientResponseError as e:
            e: aiohttp.ClientResponseError
            if e.status == 404:
                return None
            raise e
        return device

    async def create_device(self) -> dict:
        device = dict(
            deviceName=f'{self._device_name}-client',
            deviceType='DESKTOP',
            localizedModel='python',
            model='python',
            name=f'{self._device_name}',
            systemName=f'{self._device_name}',
            systemVersion='0.1'
        )

        device = await self.post(url=WDM_DEVICES, json=device)
        return device

    async def get_message(self, message_id: str) -> Optional[webexteamssdk.Message]:
        try:
            r = await self.get(url=f'https://api.ciscospark.com/v1/messages/{message_id}')
            return webexteamssdk.Message(r)
        except Exception:
            return None

    async def process_message(self, message: webexteamssdk.Message):
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(None, self._message_callback, message)

    def run(self):
        async def process(message: aiohttp.WSMessage, ignore_emails: List[str]) -> None:
            """
            Get details of message references in a given activity and call the defined callback w/ the detailed message
            data this is run in a thread to avoid blocking asynchronous handling
            :param message: websocket message to process
            :param ignore_emails: list of emails to ignore
            """
            if self._message_callback is None:
                # nothing to do if there is no callback
                return
            data = json.loads(message.data.decode('utf8'))
            data = data['data']
            if data['eventType'] != 'conversation.activity':
                return
            activity = data['activity']
            if activity['verb'] not in ['post', 'share']:
                return
            email = activity['actor']['emailAddress']
            if email in ignore_emails:
                log.debug(f'ignoring message from self')
                return
            if self._allowed_emails and email not in self._allowed_emails:
                log.debug(f'{email} not in list of allowed emails')
                return

            message_id = activity['id']
            message = await self.get_message(message_id=message_id)
            if message is None:
                return
            log.debug(f'Message from: {message.personEmail}')
            await self.process_message(message)
            return

        async def as_run():
            """
            find/create device registration and listen for messages on websocket. For posted messages a task is
            scheduled to call the configured callback with the details of the posted message. This call is executed
            in a thread so that blocking i/o in the callback does not block asynchronous handling of further messages
            received on the websocket
            """
            self._session = aiohttp.ClientSession()
            while True:
                # find/create device registration
                device = await self.find_device()
                if device:
                    log.debug('using existing device')
                else:
                    log.debug('Creating new device')
                    device = await self.create_device()

                # we need to ignore messages from our own email addresses
                me = await self.get(url='https://api.ciscospark.com/v1/people/me')
                ignore_emails = me['emails']

                wss_url = device['webSocketUrl']
                log.debug(f'WSS url: {wss_url}')
                async with self._session.ws_connect(url=wss_url, headers={'Authorization': self.auth}) as wss:
                    async for message in wss:
                        log.debug(f'got message from websocket: {message}')
                        asyncio.create_task(process(message, ignore_emails))
                    # async for
                # async with
            # while True

        # run async code
        asyncio.run(as_run())
