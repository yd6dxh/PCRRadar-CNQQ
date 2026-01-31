from msgpack import packb, unpackb
from .aiorequests import post
from random import randint
from json import loads
from hashlib import md5
from Crypto.Cipher import AES
from base64 import b64encode, b64decode
from .bsgamesdk import login, captch
from asyncio import sleep
from re import search
from datetime import datetime
from dateutil.parser import parse
from os.path import dirname, join, exists
import traceback
import uuid

def gen_device_id():
    return uuid.uuid4().hex
# 你当前实际在用的
apiroot = 'https://le1-prod-all-gs-gzlj.bilibiligame.net'

curpath = dirname(__file__)
config = join(curpath, 'version.txt')

# 没有 version.txt 也能跑；一旦解析到新版本会自动创建
version = "4.9.4"
if exists(config):
    with open(config, encoding='utf-8') as fp:
        version = fp.read().strip()

defaultHeaders = {
    'Accept-Encoding': 'gzip',
    'User-Agent': 'Dalvik/2.1.0 (Linux, U, Android 5.1.1, PCRT00 Build/LMY48Z)',
    'X-Unity-Version': '2018.4.30f1',
    'APP-VER': version,
    'BATTLE-LOGIC-VERSION': '4',
    'BUNDLE-VER': '',
    'DEVICE': '2',
    'DEVICE-ID': gen_device_id(),
    'DEVICE-NAME': 'OPPO PCRT00',
    'EXCEL-VER': '1.0.0',
    'GRAPHICS-DEVICE-NAME': 'Adreno (TM) 640',
    'IP-ADDRESS': '10.0.2.15',
    'KEYCHAIN': '',
    'LOCALE': 'CN',
    'PLATFORM-OS-VERSION': 'Android OS 5.1.1 / API-22 (LMY48Z/rel.se.infra.20200612.100533)',
    'REGION-CODE': '',
    'RES-KEY': 'ab00a0a6dd915a052a2ef7fd649083e5',
    'RES-VER': '10002200',
    'SHORT-UDID': '0'
}

# === 按 apiclient：CBC IV 用这一组 ===
_AES_IV = b'7Fk9Lm3Np8Qr4Sv2'


class ApiException(Exception):
    def __init__(self, message, code):
        super().__init__(message)
        self.code = code


class bsdkclient:
    """
    acccountinfo = {
        'account': '',
        'password': '',
        'platform': 2,
        'channel': 1,
    }
    """
    def __init__(self, acccountinfo, captchaVerifier, errlogger):
        self.account = acccountinfo['account']
        self.pwd = acccountinfo['password']
        self.platform = acccountinfo['platform']
        self.channel = acccountinfo['channel']
        self.captchaVerifier = captchaVerifier
        self.errlogger = errlogger

    async def login(self):
        while True:
            resp = await login(self.account, self.pwd, self.captchaVerifier)
            if resp['code'] == 0:
                break
            await self.errlogger(resp['message'])
        return resp['uid'], resp['access_key']


class pcrclient:
    def __init__(self, bsclient: bsdkclient):
        self.viewer_id = 0
        self.bsdk = bsclient
        self.errlogger = getattr(bsclient, "errlogger", None)

        self.headers = {k: defaultHeaders[k] for k in defaultHeaders.keys()}

        self.shouldLogin = True
        self.shouldLoginB = True

    async def bililogin(self):
        self.uid, self.access_key = await self.bsdk.login()
        self.platform = self.bsdk.platform
        self.channel = self.bsdk.channel
        self.headers['PLATFORM'] = str(self.platform)
        self.headers['PLATFORM-ID'] = str(self.platform)
        self.headers['CHANNEL-ID'] = str(self.channel)
        self.shouldLoginB = False

    @staticmethod
    def createkey() -> bytes:
        return bytes([ord('0123456789abcdef'[randint(0, 15)]) for _ in range(32)])

    @staticmethod
    def add_to_16(b: bytes) -> bytes:
        n = len(b) % 16
        n = n // 16 * 16 - n + 16
        return b + (n * bytes([n]))

    @staticmethod
    def pack(data: object, key: bytes) -> bytes:
        aes = AES.new(key, AES.MODE_CBC, _AES_IV)
        return aes.encrypt(pcrclient.add_to_16(packb(data, use_bin_type=False))) + key

    @staticmethod
    def encrypt(data: str, key: bytes) -> bytes:
        aes = AES.new(key, AES.MODE_CBC, _AES_IV)
        return aes.encrypt(pcrclient.add_to_16(data.encode('utf8'))) + key

    @staticmethod
    def decrypt(data: bytes):
        data = b64decode(data.decode('utf8'))
        aes = AES.new(data[-32:], AES.MODE_CBC, _AES_IV)
        return aes.decrypt(data[:-32]), data[-32:]

    @staticmethod
    def unpack(data: bytes):
        data = b64decode(data.decode('utf8'))
        aes = AES.new(data[-32:], AES.MODE_CBC, _AES_IV)
        dec = aes.decrypt(data[:-32])
        return unpackb(dec[:-dec[-1]], strict_map_key=False), data[-32:]

    async def callapi(self, apiurl: str, request: dict, crypted: bool = True, noerr: bool = False):
        key = pcrclient.createkey()
        try:
            if self.viewer_id is not None:
                request['viewer_id'] = b64encode(
                    pcrclient.encrypt(str(self.viewer_id), key)
                ) if crypted else str(self.viewer_id)

            raw = await (await post(
                apiroot + apiurl,
                data=pcrclient.pack(request, key) if crypted else str(request).encode('utf8'),
                headers=self.headers,
                timeout=10
            )).content

            response = pcrclient.unpack(raw)[0] if crypted else loads(raw)
            data_headers = response['data_headers']

            # === 完全按 apiclient：在 maintenance_status 的 data_headers.store_url 触发版本更新 ===
            if apiurl == "/source_ini/get_maintenance_status?format=json" and "store_url" in data_headers:
                m = search(r'(?<=gzlj_)(\d+\.\d+\.\d+)', data_headers["store_url"])
                if not m:
                    raise ValueError("无法解析版本号，请手动修改version")
                new_ver = m.group(1)

                global version
                version = new_ver
                defaultHeaders['APP-VER'] = new_ver
                self.headers['APP-VER'] = new_ver
                with open(config, "w", encoding="utf-8") as fp:
                    fp.write(new_ver)

                raise ApiException(f"版本已更新:{new_ver}", 0)

            if 'sid' in data_headers and data_headers["sid"] != '':
                t = md5()
                t.update((data_headers['sid'] + 'c!SID!n').encode('utf8'))
                self.headers['SID'] = t.hexdigest()

            if 'request_id' in data_headers:
                self.headers['REQUEST-ID'] = data_headers['request_id']

            if 'viewer_id' in data_headers:
                self.viewer_id = data_headers['viewer_id']

            data = response['data']
            if not noerr and 'server_error' in data:
                err = data['server_error']
                if self.errlogger:
                    try:
                        await self.errlogger(f"pcrclient: {apiurl} api failed {err}")
                    except:
                        pass
                if "store_url" in data_headers:
                    raise ApiException(f"版本自动更新失败：({err['message']})", err['status'])
                raise ApiException(err['message'], err['status'])

            print(f'pcrclient: {apiurl} api called')
            return data

        except Exception as e:
            if self.errlogger:
                try:
                    await self.errlogger(f"pcrclient: exception at {apiurl}: {repr(e)}")
                    await self.errlogger(traceback.format_exc())
                except:
                    pass
            self.shouldLogin = True
            raise

    async def login(self):
        if self.shouldLoginB:
            await self.bililogin()

        if 'REQUEST-ID' in self.headers:
            self.headers.pop('REQUEST-ID')

        while True:
            manifest = await self.callapi('/source_ini/get_maintenance_status?format=json', {}, False, noerr=True)
            if 'maintenance_message' not in manifest:
                break

            try:
                match = search(r'\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d', manifest['maintenance_message']).group()
                end = parse(match)
                print(f'server is in maintenance until {match}')
                while datetime.now() < end:
                    await sleep(1)
            except:
                print('server is in maintenance. waiting for 60 secs')
                await sleep(60)

        ver = manifest['required_manifest_ver']
        print(f'using manifest ver = {ver}')
        self.headers['MANIFEST-VER'] = str(ver)

        lres = await self.callapi('/tool/sdk_login', {
            'uid': str(self.uid),
            'access_key': self.access_key,
            'channel': str(self.channel),
            'platform': str(self.platform)
        })

        retry_times = 0
        while retry_times < 5:
            retry_times += 1
            if "is_risk" in lres and lres["is_risk"] == 1:
                print(lres)
                while True:
                    try:
                        cap = await captch()
                        challenge, gt_user_id, validate = await self.bsdk.captchaVerifier(
                            cap['gt'], cap['challenge'], cap['gt_user_id']
                        )
                        if validate:
                            lres = await self.callapi(
                                "/tool/sdk_login",
                                {
                                    "uid": str(self.uid),
                                    "access_key": self.access_key,
                                    "channel": str(self.channel),
                                    "platform": str(self.platform),
                                    'challenge': challenge,
                                    'validate': validate,
                                    'seccode': validate + "|jordan",
                                    'captcha_type': '1',
                                    'image_token': '',
                                    'captcha_code': '',
                                },
                            )
                            break
                    except:
                        self.shouldLoginB = True
                        await self.bililogin()
            else:
                break
        else:
            raise Exception("验证码错误")

        gamestart = await self.callapi('/check/game_start', {
            'apptype': 0,
            'campaign_data': '',
            'campaign_user': randint(0, 99999)
        })

        if not gamestart['now_tutorial']:
            raise Exception("该账号没过完教程!")

        await self.callapi('/load/index', {'carrier': 'OPPO'})
        await self.callapi('/home/index', {
            'message_id': 1,
            'tips_id_list': [],
            'is_first': 1,
            'gold_history': 0
        })

        self.shouldLogin = False
