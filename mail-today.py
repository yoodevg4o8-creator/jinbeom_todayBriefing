#!/usr/bin/env python3
"""오늘의 아침 브리핑 한 장: 날씨 + 법령 개정 + 하이웍스 오늘 메일 (읽기 전용, 파이썬 표준 라이브러리만 사용)

사용법
  python3 mail-today.py --init     설정 파일 틀을 만듭니다 (account.local.json, settings.local.json)
  python3 mail-today.py --check    하이웍스 로그인만 확인합니다 (메일을 읽지 않음)
  python3 mail-today.py            브리핑 페이지를 만듭니다 → out/mail-today.html
  python3 mail-today.py --demo     가짜 자료로 화면 모양만 확인합니다 (메일 서버 접속 없음)
  python3 mail-today.py --send 초안.json                 보낼 메일을 미리보기만 합니다 (보내지 않음, 확인코드 출력)
  python3 mail-today.py --send 초안.json --confirm 코드   미리보기와 같은 내용일 때만 실제로 보냅니다
  (보내기는 settings.local.json 에 "allowSend": true 를 적어야 켜집니다. 기본은 꺼짐)

안전 원칙
  - 메일은 읽기만 합니다. 삭제(DELE)·발송·이동은 코드에 없습니다.
  - 비밀번호는 account.local.json 에만 적고, 화면·로그·결과 파일에 쓰지 않습니다.
  - 하이웍스 웹메일 설정에서 "POP3로 가져간 메일을 읽지 않음으로 표시"를 켜 두면 읽음 처리도 되지 않습니다.
"""
import base64
import datetime
import email
import email.policy
import email.utils
import hashlib
import html
import json
import mimetypes
import os
import poplib
import re
import smtplib
import ssl
import stat
import sys
from email.message import EmailMessage
import urllib.error
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ACCOUNT_FILE = os.path.join(HERE, 'account.local.json')
SETTINGS_FILE = os.path.join(HERE, 'settings.local.json')
STATE_FILE = os.path.join(HERE, 'state.local.json')
OUT_DIR = os.path.join(HERE, 'out')
KST = datetime.timezone(datetime.timedelta(hours=9))
SECRET_SUBJECT = re.compile(r'verification code|one.time|인증\s*(코드|번호)|비밀번호|password reset', re.I)
DOC_EXT = ('.pdf', '.docx', '.xlsx', '.pptx', '.hwp', '.hwpx', '.txt', '.csv', '.png', '.jpg', '.jpeg')
DEFAULT_LAWS = ['산업안전보건법', '산업안전보건법 시행령', '산업안전보건법 시행규칙', '산업안전보건기준에 관한 규칙']
MAX_MAILS_SCANNED = 80          # 받은편지함 맨 뒤(최신)에서 이만큼까지만 확인합니다
MAX_MAIL_BYTES = 25 * 1024 * 1024  # 이보다 큰 메일은 내려받지 않습니다
EMBED_FILE_LIMIT = 3 * 1024 * 1024   # 첨부를 페이지 안에 넣는 파일당 한도
EMBED_TOTAL_LIMIT = 20 * 1024 * 1024  # 페이지 전체 한도


# ───────────────────────── 설정 ─────────────────────────
def read_json(path, default=None):
    try:
        with open(path, encoding='utf-8') as handle:
            return json.load(handle)
    except FileNotFoundError:
        return default
    except (OSError, ValueError):
        raise SystemExit('설정 파일을 읽지 못했습니다: ' + os.path.basename(path) + ' (JSON 형식을 확인해 주세요)')


def write_private(path, text):
    # 소유자만 읽고 쓸 수 있게(0600) 새로 만듭니다.
    flags = os.O_WRONLY | os.O_CREAT | os.O_TRUNC
    fd = os.open(path, flags, 0o600)
    with os.fdopen(fd, 'w', encoding='utf-8') as handle:
        handle.write(text)
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


def init_files():
    created = []
    if not os.path.exists(ACCOUNT_FILE):
        write_private(ACCOUNT_FILE, json.dumps({
            'account': '여기에 하이웍스 아이디(메일 주소)를 적으세요',
            'password': '여기에 비밀번호를 적으세요 (OTP를 쓰면 메일 전용 비밀번호)'
        }, ensure_ascii=False, indent=2) + '\n')
        created.append('account.local.json')
    if not os.path.exists(SETTINGS_FILE):
        write_private(SETTINGS_FILE, json.dumps({
            'exclude': [],
            'mailFrom': '오늘 00:00',
            'mailTo': '지금',
            'allowSend': False,
            'keepDays': 7,
            'weather': {'name': '서울', 'latitude': 37.5665, 'longitude': 126.978},
            'lawWatch': True,
            'laws': DEFAULT_LAWS,
            'lawOc': '',
            'lawDays': 14
        }, ensure_ascii=False, indent=2) + '\n')
        created.append('settings.local.json')
    print(json.dumps({'created': created, 'next': 'account.local.json 을 직접 열어 아이디와 비밀번호를 적으세요'}, ensure_ascii=False))


def load_account():
    data = read_json(ACCOUNT_FILE)
    if data is None and os.environ.get('HIWORKS_MAIL_ACCOUNT') and os.environ.get('HIWORKS_MAIL_PASSWORD'):
        # 파일 대신 환경 변수로 줄 수도 있습니다. (개인 AI 비서의 .env.local 을 그대로 쓰는 경우)
        data = {'account': os.environ['HIWORKS_MAIL_ACCOUNT'], 'password': os.environ['HIWORKS_MAIL_PASSWORD']}
    if not isinstance(data, dict) or not isinstance(data.get('account'), str) or not isinstance(data.get('password'), str):
        raise SystemExit('account.local.json 이 없거나 형식이 맞지 않습니다. 먼저 --init 을 실행하고 아이디와 비밀번호를 적어 주세요.')
    account, password = data['account'].strip(), data['password']
    if '@' not in account or password.startswith('여기에'):
        raise SystemExit('account.local.json 에 아이디(메일 주소)와 비밀번호를 적어 주세요.')
    return account, password


def permission_warning():
    # 다른 사용자가 읽을 수 있는 권한이면 알려 줍니다. (윈도우는 건너뜁니다)
    try:
        mode = os.stat(ACCOUNT_FILE).st_mode
        if mode & (stat.S_IRWXG | stat.S_IRWXO):
            return 'account.local.json 파일 권한이 넓습니다. chmod 600 account.local.json 을 실행해 주세요.'
    except OSError:
        pass
    return ''


# ───────────────────────── 메일 (하이웍스 POP3, 읽기 전용) ─────────────────────────
def html_to_text(text):
    text = re.sub(r'(?is)<(style|script|head)\b.*?</\1>', ' ', text)
    text = re.sub(r'(?i)<br\s*/?>|</(p|div|tr|li|h[1-6])>', '\n', text)
    return html.unescape(re.sub(r'<[^>]+>', ' ', text))


def tidy(body):
    return re.sub(r'[ \t]+', ' ', re.sub(r'\n\s*\n+', '\n\n', body)).strip()


def parse_when(value):
    try:
        when = email.utils.parsedate_to_datetime(str(value))
        if when.tzinfo is None:
            when = when.replace(tzinfo=datetime.timezone.utc)
        return when.astimezone(KST)
    except (TypeError, ValueError):
        return None


def save_attachments(msg, folder):
    """문서·그림 파일만 저장합니다. 실행 파일은 받지 않고 파일 이름은 안전한 글자만 남깁니다."""
    names, saved = [], []
    for attachment in msg.iter_attachments():
        filename = attachment.get_filename() or '(이름 없음)'
        names.append(filename)
        ext = os.path.splitext(filename)[1].lower()
        if ext not in DOC_EXT or len(saved) >= 10:
            continue
        try:
            data = attachment.get_content()
            if isinstance(data, str):
                data = data.encode('utf-8')
            if not isinstance(data, bytes) or len(data) > 15 * 1024 * 1024:
                continue
            os.makedirs(folder, exist_ok=True)
            safe = re.sub(r'[^0-9A-Za-z가-힣._() -]', '_', os.path.basename(filename))[:80] or ('file' + ext)
            path = os.path.join(folder, safe)
            with open(path, 'wb') as handle:
                handle.write(data)
            os.chmod(path, 0o600)
            saved.append(path)
        except Exception:
            continue
    return names[:20], saved


def read_message(raw, folder):
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    try:
        part = msg.get_body(preferencelist=('plain', 'html'))
        text = part.get_content() if part is not None else ''
        body = html_to_text(text) if part is not None and part.get_content_type() == 'text/html' else text
    except Exception:
        body = ''
    names, saved = save_attachments(msg, folder)
    return {'from': str(msg.get('From', ''))[:300], 'replyTo': str(msg.get('Reply-To', ''))[:300],
            'messageId': str(msg.get('Message-ID', '')).strip()[:300], 'references': ' '.join(str(msg.get('References', '')).split())[:2000],
            'to': str(msg.get('To', ''))[:500], 'cc': str(msg.get('Cc', ''))[:500],
            'subject': str(msg.get('Subject', ''))[:300], 'date': str(msg.get('Date', ''))[:100],
            'body': tidy(body)[:20000], 'attachments': names, 'savedFiles': saved}


def connect(account, password):
    client = poplib.POP3_SSL('pop3s.hiworks.com', 995, timeout=25, context=ssl.create_default_context())
    client.user(account)
    client.pass_(password)
    return client


DAYS = {'오늘': 0, 'today': 0, '어제': 1, 'yesterday': 1, '그제': 2, '그저께': 2}


def parse_point(text, now, default_time):
    """"어제 18:00", "오늘 06:55", "지금" 같은 글을 시각으로 바꿉니다."""
    m = re.fullmatch(r'\s*(오늘|today|어제|yesterday|그제|그저께|지금|now)(?:\s+(\d{1,2}):(\d{2}))?\s*', str(text), re.I)
    if not m:
        raise SystemExit('settings.local.json 의 mailFrom/mailTo 형식을 확인해 주세요. 예: "어제 18:00", "오늘 06:55", "지금"')
    word = m.group(1).lower()
    if word in ('지금', 'now'):
        return now
    hour, minute = (int(m.group(2)), int(m.group(3))) if m.group(2) else default_time
    if hour > 23 or minute > 59:
        raise SystemExit('settings.local.json 의 시각은 00:00~23:59 사이로 적어 주세요.')
    day = (now - datetime.timedelta(days=DAYS[word])).replace(hour=hour, minute=minute, second=0, microsecond=0)
    return day


def mail_window(settings, now=None):
    """메일을 가져올 구간입니다. 기본은 오늘 00:00 ~ 지금. 설정: mailFrom, mailTo"""
    now = now or datetime.datetime.now(KST)
    start = parse_point(settings.get('mailFrom') or '오늘 00:00', now, (0, 0))
    end = parse_point(settings.get('mailTo') or '지금', now, (23, 59))
    # 월요일 아침에 "어제"로 시작하면 주말 동안 쌓인 메일이 빠지므로, 시작을 금요일(같은 시각)로 당깁니다.
    if now.weekday() == 0 and re.match(r'\s*(어제|yesterday)\b', str(settings.get('mailFrom') or ''), re.I):
        start -= datetime.timedelta(days=2)
    if start >= end:
        raise SystemExit('mailFrom 이 mailTo 보다 늦습니다. 구간을 확인해 주세요.')
    return start, end


def read_today(account, password, exclude, attach_dir, window=None):
    patterns = [re.compile(p, re.I) for p in exclude if isinstance(p, str) and p][:50]
    start, end = window or mail_window({})
    client = connect(account, password)
    try:
        uids = client.uidl()[1]
        sizes = {}
        for row in client.list()[1]:
            parts = row.decode('ascii', errors='replace').split()
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                sizes[int(parts[0])] = int(parts[1])
        messages, skipped, reached_older, later = [], 0, False, 0
        scanned = uids[-MAX_MAILS_SCANNED:]
        for row in reversed(scanned):
            number = int(row.decode('ascii', errors='replace').split(' ', 1)[0])
            head = email.message_from_bytes(b'\r\n'.join(client.top(number, 0)[1]), policy=email.policy.default)
            when = parse_when(head.get('Date', ''))
            if when is not None and when < start:
                reached_older = True
                continue
            if when is not None and when > end:
                later += 1  # 구간 뒤에 도착한 메일은 가져오지 않습니다
                continue
            sender, subject = str(head.get('From', '')), str(head.get('Subject', ''))
            if any(p.search(sender) or p.search(subject) for p in patterns) or SECRET_SUBJECT.search(subject):
                skipped += 1  # 제외 대상은 본문을 내려받지 않습니다
                continue
            if sizes.get(number, 0) > MAX_MAIL_BYTES:
                messages.append({'from': sender[:300], 'to': '', 'cc': '', 'subject': subject[:300], 'date': str(head.get('Date', ''))[:100],
                                 'body': '(메일 용량이 너무 커서 내려받지 않았습니다. 웹메일에서 확인하세요)', 'attachments': [], 'savedFiles': []})
                continue
            folder = os.path.join(attach_dir, str(len(messages) + 1))
            messages.append(read_message(b'\r\n'.join(client.retr(number)[1]), folder))
        limited = len(uids) > MAX_MAILS_SCANNED and not reached_older
        return {'messages': messages, 'skipped': skipped, 'limited': limited, 'later': later, 'start': start, 'end': end}
    finally:
        try:
            client.quit()
        except Exception:
            pass


def check_login(account, password):
    client = connect(account, password)
    try:
        client.stat()
    finally:
        try:
            client.quit()
        except Exception:
            pass


# ───────────────────────── 날씨 (Open-Meteo, 키 없음) ─────────────────────────
WEATHER = {0: ('맑음', '☀️'), 1: ('대체로 맑음', '🌤️'), 2: ('구름 조금', '⛅'), 3: ('흐림', '☁️'), 45: ('안개', '🌫️'), 48: ('짙은 안개', '🌫️'),
           51: ('약한 이슬비', '🌦️'), 53: ('이슬비', '🌦️'), 55: ('강한 이슬비', '🌧️'), 56: ('어는 이슬비', '🌧️'), 57: ('강한 어는 이슬비', '🌧️'),
           61: ('약한 비', '🌦️'), 63: ('비', '🌧️'), 65: ('강한 비', '🌧️'), 66: ('어는 비', '🌧️'), 67: ('강한 어는 비', '🌧️'),
           71: ('약한 눈', '🌨️'), 73: ('눈', '🌨️'), 75: ('강한 눈', '❄️'), 77: ('싸락눈', '🌨️'), 80: ('약한 소나기', '🌦️'), 81: ('소나기', '🌧️'),
           82: ('강한 소나기', '⛈️'), 85: ('약한 눈 소나기', '🌨️'), 86: ('강한 눈 소나기', '❄️'), 95: ('뇌우', '⛈️'),
           96: ('우박 동반 뇌우', '⛈️'), 99: ('강한 우박 동반 뇌우', '⛈️')}


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def fetch_json(url, limit=6 * 1024 * 1024, timeout=15):
    opener = urllib.request.build_opener(NoRedirect)
    request = urllib.request.Request(url, headers={'Accept': 'application/json', 'User-Agent': 'mail-today/1.0'})
    with opener.open(request, timeout=timeout) as response:
        raw = response.read(limit + 1)
    if len(raw) > limit:
        raise ValueError('too large')
    return json.loads(raw.decode('utf-8'))


def number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)


def get_weather(setting):
    if setting is False:
        return None
    place = setting if isinstance(setting, dict) else {}
    lat, lon = number(place.get('latitude', 37.5665)), number(place.get('longitude', 126.978))
    if lat is None or lon is None or not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return {'error': 'settings.local.json 의 weather 위도·경도를 확인해 주세요.'}
    name = str(place.get('name') or '서울')[:40]
    query = urllib.parse.urlencode({
        'latitude': round(lat, 2), 'longitude': round(lon, 2),  # 약 1km 단위로만 보냅니다
        'current': 'temperature_2m,apparent_temperature,relative_humidity_2m,weather_code,wind_speed_10m',
        'daily': 'temperature_2m_max,temperature_2m_min,precipitation_probability_max',
        'timezone': 'auto', 'forecast_days': 1, 'wind_speed_unit': 'ms'})
    try:
        data = fetch_json('https://api.open-meteo.com/v1/forecast?' + query)
        cur, day = data.get('current') or {}, data.get('daily') or {}
        code = int(cur['weather_code'])
        label, icon = WEATHER.get(code, ('날씨 정보 없음', '🌡️'))
        first = lambda key: (day.get(key) or [None])[0]
        return {'name': name, 'label': label, 'icon': icon, 'temp': number(cur.get('temperature_2m')),
                'feels': number(cur.get('apparent_temperature')), 'humidity': number(cur.get('relative_humidity_2m')),
                'wind': number(cur.get('wind_speed_10m')), 'max': number(first('temperature_2m_max')),
                'min': number(first('temperature_2m_min')), 'rain': number(first('precipitation_probability_max'))}
    except Exception:
        return {'error': '날씨를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.'}


# ───────────────────────── 법령 개정 (국가법령정보센터 신구조문 대비표) ─────────────────────────
def as_list(value):
    return [] if value is None else value if isinstance(value, list) else [value]


def plain(value):
    return re.sub(r'\s+', ' ', re.sub(r'</?[A-Za-z][^>]*>', '', str(value if value is not None else ''))).strip()


def clip(value, limit=220):
    return value if len(value) <= limit else value[:limit] + '…'


def compact(value):
    return re.sub(r'\s+', '', plain(value))


def same_mark(value):
    return re.search(r'\((현행과같음|생략)\)', compact(value)) is not None


def changed_articles(service):
    """대비표에서 바뀐 조문 줄만 뽑습니다. 구조문과 신조문은 같은 순서로 나란히 옵니다."""
    news = as_list((service or {}).get('신조문목록', {}).get('조문'))
    olds = as_list((service or {}).get('구조문목록', {}).get('조문'))
    out, heading = [], ''
    for i in range(max(len(news), len(olds))):
        n = plain(news[i].get('content') if i < len(news) and isinstance(news[i], dict) else '')
        o = plain(olds[i].get('content') if i < len(olds) and isinstance(olds[i], dict) else '')
        head = re.match(r'^(제\d+조(?:의\d+)?)(?:\(([^)]*)\))?', n or o)
        if head:
            heading = head.group(1) + ('(' + head.group(2) + ')' if head.group(2) else '')
        if compact(n) == compact(o) or same_mark(n) or same_mark(o):
            continue
        out.append({'article': heading, 'old': clip(o) or '(없음)', 'new': clip(n) or '(삭제)'})
    return out


class ApiRejected(Exception):
    pass


def ymd(value):
    value = str(value)
    return value[:4] + '-' + value[4:6] + '-' + value[6:] if re.fullmatch(r'\d{8}', value) else ''


def get_law(settings, state, offline_rows=None):
    """새로 공포된 개정만 보여 줍니다. 읽기만 하며 법적 판단을 대신하지 않습니다."""
    if settings.get('lawWatch') is False:
        return None
    laws = [x for x in as_list(settings.get('laws') or DEFAULT_LAWS) if isinstance(x, str) and x.strip()][:10]
    oc = str(settings.get('lawOc') or '').strip()
    if not oc:
        return {'error': '법령 개정 확인을 쓰려면 국가법령정보 공동활용(open.law.go.kr)에서 인증값(OC)을 발급받아 settings.local.json 의 lawOc 에 적어 주세요. 쓰지 않으려면 lawWatch 를 false 로 바꾸세요.'}
    if not re.fullmatch(r'[A-Za-z0-9_.@+\-]{1,200}', oc):
        return {'error': 'settings.local.json 의 lawOc 값을 확인해 주세요.'}
    days = settings.get('lawDays') if isinstance(settings.get('lawDays'), int) and 1 <= settings.get('lawDays') <= 90 else 14
    since = (datetime.datetime.now(KST) - datetime.timedelta(days=days)).strftime('%Y%m%d')
    seen = {k: set(map(str, v)) for k, v in (state.get('lawSeen') or {}).items() if isinstance(v, list)}
    blocks, new_seen = [], {}

    def call(route, params):
        query = urllib.parse.urlencode(dict({'OC': oc, 'type': 'JSON'}, **params))
        return fetch_json('https://www.law.go.kr/DRF/' + route + '?' + query)

    try:
        for name in laws:
            data = call('lawSearch.do', {'target': 'oldAndNew', 'query': name.replace(' ', ''), 'display': '20'})
            if not isinstance(data, dict) or 'OldAndNewLawSearch' not in data:
                # 인증 실패 등 서버가 오류를 알려 준 경우입니다. "개정 없음"으로 보이면 안 됩니다.
                raise ApiRejected(str(data.get('result', ''))[:100] if isinstance(data, dict) else '')
            rows = [r for r in as_list((data.get('OldAndNewLawSearch') or {}).get('oldAndNew'))
                    if isinstance(r, dict) and str(r.get('신구법명', '')).strip() == name and re.fullmatch(r'\d{1,10}', str(r.get('신구법일련번호', '')))]
            known = seen.get(name, set())
            for row in [r for r in rows if str(r.get('공포일자', '')) >= since][:2]:
                key = str(row['신구법일련번호'])
                service = (call('lawService.do', {'target': 'oldAndNew', 'MST': key}) or {}).get('OldAndNewService') or {}
                changes = changed_articles(service)
                blocks.append({'name': name, 'kind': str(row.get('제개정구분명', '')).strip(), 'promulgated': ymd(row.get('공포일자')),
                               'effective': ymd((service.get('신조문_기본정보') or {}).get('시행일자') or row.get('시행일자')),
                               'count': len(changes), 'changes': changes[:8], 'more': max(0, len(changes) - 8), 'fresh': key not in known})
            new_seen[name] = sorted(known | {str(r['신구법일련번호']) for r in rows})[-50:]
    except ApiRejected as rejected:
        return {'error': '국가법령정보 서버가 요청을 거부했습니다' + (' (' + str(rejected) + ')' if str(rejected) else '') + '. 인증값(lawOc)과 이용 신청·서버 IP 등록 상태를 확인해 주세요.'}
    except Exception:
        return {'error': '법령 개정 정보를 읽지 못했습니다. 인터넷 연결을 확인해 주세요.'}
    state['lawSeen'] = new_seen
    return {'blocks': blocks, 'days': days, 'laws': laws}


# ───────────────────────── 보내기 (기본 꺼짐, 미리보기 → 확인코드 → 발송) ─────────────────────────
ADDRESS = re.compile(r'[^\s<>@,;]+@[^\s<>@,;]+\.[A-Za-z]{2,}')
MAX_RECIPIENTS, MAX_ATTACH, MAX_ATTACH_BYTES, MAX_ATTACH_TOTAL, DAILY_SEND_LIMIT = 10, 3, 5 * 1024 * 1024, 10 * 1024 * 1024, 20


def clean_recipients(value, name):
    if value is None:
        value = []
    if not isinstance(value, list) or len(value) > MAX_RECIPIENTS:
        raise SystemExit('초안의 %s 는 주소 목록(최대 %d명)이어야 합니다.' % (name, MAX_RECIPIENTS))
    for item in value:
        if not isinstance(item, str) or not ADDRESS.fullmatch(item):
            raise SystemExit('초안의 %s 에 올바르지 않은 주소가 있습니다.' % name)
    return value


def load_draft(path):
    """초안 파일(JSON)을 읽고 검사합니다. {to, cc, subject, text, attachments[], inReplyTo, references}"""
    draft = read_json(path)
    if not isinstance(draft, dict):
        raise SystemExit('초안 파일을 읽지 못했습니다. JSON 형식인지 확인해 주세요.')
    to, cc = clean_recipients(draft.get('to'), 'to'), clean_recipients(draft.get('cc'), 'cc')
    if not to:
        raise SystemExit('초안에 받는 사람(to)이 없습니다.')
    subject, text = draft.get('subject'), draft.get('text')
    if not isinstance(subject, str) or not subject.strip() or re.search(r'[\r\n]', subject) or len(subject) > 300:
        raise SystemExit('초안의 제목(subject)은 한 줄, 300자 이하여야 합니다.')
    if not isinstance(text, str) or not text.strip() or len(text) > 100000:
        raise SystemExit('초안의 본문(text)이 비었거나 너무 깁니다.')
    reply, refs = draft.get('inReplyTo'), draft.get('references')
    if reply is not None and (not isinstance(reply, str) or not re.fullmatch(r'<[^<>\s]{1,300}>', reply)):
        raise SystemExit('초안의 inReplyTo 는 <...> 형태의 메시지 번호여야 합니다.')
    if refs is not None and (not isinstance(refs, str) or not re.fullmatch(r'(<[^<>\s]{1,300}>\s*){1,20}', refs)):
        raise SystemExit('초안의 references 형식을 확인해 주세요.')
    files, total = [], 0
    for item in as_list(draft.get('attachments')):
        if not isinstance(item, str) or len(files) >= MAX_ATTACH:
            raise SystemExit('첨부는 파일 경로 문자열로 최대 %d개까지 가능합니다.' % MAX_ATTACH)
        name = os.path.basename(item)
        if re.search(r'[\r\n/\\]', name) or not os.path.isfile(item):
            raise SystemExit('첨부 파일을 찾지 못했거나 이름이 올바르지 않습니다: ' + name)
        with open(item, 'rb') as handle:
            data = handle.read(MAX_ATTACH_BYTES + 1)
        total += len(data)
        if len(data) > MAX_ATTACH_BYTES or total > MAX_ATTACH_TOTAL:
            raise SystemExit('첨부가 너무 큽니다. (파일당 5MB, 전체 10MB 이하)')
        files.append({'name': name, 'data': data, 'sha256': hashlib.sha256(data).hexdigest()})
    return {'to': to, 'cc': cc, 'subject': subject.strip(), 'text': text, 'inReplyTo': reply, 'references': (refs or '').strip(),
            'files': files}


def draft_code(draft, account):
    """미리보기와 실제 발송이 같은 내용인지 확인하는 코드입니다. 내용이 한 글자라도 바뀌면 달라집니다."""
    canon = json.dumps({'from': account, 'to': draft['to'], 'cc': draft['cc'], 'subject': draft['subject'], 'text': draft['text'],
                        'inReplyTo': draft['inReplyTo'], 'references': draft['references'],
                        'files': [[f['name'], f['sha256']] for f in draft['files']]}, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(canon.encode('utf-8')).hexdigest()[:10]


def preview_text(draft, account, code):
    files = ', '.join('%s (%d KB)' % (f['name'], (len(f['data']) + 1023) // 1024) for f in draft['files']) or '없음'
    return '\n'.join([
        '── 보낼 메일 미리보기 (아직 보내지 않았습니다) ──',
        '보내는 사람: ' + account, '받는 사람: ' + ', '.join(draft['to']), '참조: ' + (', '.join(draft['cc']) or '없음'),
        '제목: ' + draft['subject'], '첨부: ' + files, '이어 붙이기(답장): ' + ('예' if draft['inReplyTo'] else '아니오'),
        '── 본문 ──', draft['text'], '──────────',
        '확인코드: ' + code, '보내려면 같은 초안으로 다시 실행: python3 mail-today.py --send <초안파일> --confirm ' + code])


def build_message(account, draft):
    msg = EmailMessage()
    msg['From'], msg['To'], msg['Subject'] = account, ', '.join(draft['to']), draft['subject']
    if draft['cc']:
        msg['Cc'] = ', '.join(draft['cc'])
    msg['Date'] = email.utils.formatdate(localtime=True)
    msg['Message-ID'] = email.utils.make_msgid()
    if draft['inReplyTo']:
        msg['In-Reply-To'] = draft['inReplyTo']
        msg['References'] = (draft['references'] + ' ' + draft['inReplyTo']).strip() if draft['inReplyTo'] not in draft['references'] else draft['references']
    msg.set_content(draft['text'])
    for item in draft['files']:
        msg.add_attachment(item['data'], maintype='application', subtype='octet-stream', filename=item['name'])
    return msg


def send_flow(path, confirm, settings, state):
    """미리보기만(confirm 없음) 하거나, 코드가 맞을 때 한 번 보냅니다. 실패해도 자동으로 다시 보내지 않습니다."""
    if settings.get('allowSend') is not True:
        raise SystemExit('보내기가 꺼져 있습니다. 쓰려면 settings.local.json 에 "allowSend": true 를 적어 주세요.')
    account, password = load_account()
    draft = load_draft(path)
    code = draft_code(draft, account)
    if not confirm:
        print(preview_text(draft, account, code))
        return
    if confirm != code:
        raise SystemExit('확인코드가 다릅니다. 초안이 미리보기 뒤에 바뀌었을 수 있습니다. 다시 미리보기를 확인해 주세요. (보내지 않았습니다)')
    today = datetime.datetime.now(KST).strftime('%Y-%m-%d')
    sent = state.get('sent') if isinstance(state.get('sent'), dict) and state['sent'].get('date') == today else {'date': today, 'count': 0, 'codes': []}
    if code in sent['codes']:
        raise SystemExit('같은 내용의 메일을 오늘 이미 보냈습니다. 중복 발송을 막기 위해 보내지 않았습니다.')
    if sent['count'] >= DAILY_SEND_LIMIT:
        raise SystemExit('오늘 보낼 수 있는 한도(%d통)를 넘었습니다.' % DAILY_SEND_LIMIT)
    # 보내기 직전에 기록부터 남깁니다. 도중에 실패해도 자동으로 다시 보내지 않기 위해서입니다.
    sent['count'] += 1
    sent['codes'].append(code)
    state['sent'] = sent
    write_private(STATE_FILE, json.dumps(state, ensure_ascii=False))
    try:
        with smtplib.SMTP_SSL('smtps.hiworks.com', 465, timeout=25, context=ssl.create_default_context()) as client:
            client.login(account, password)
            client.send_message(build_message(account, draft), from_addr=account, to_addrs=draft['to'] + draft['cc'])
    except Exception:
        raise SystemExit('메일을 보내지 못했습니다. 아이디·비밀번호, 하이웍스 SMTP 사용 설정, 인터넷 연결을 확인해 주세요. 자동으로 다시 보내지 않습니다. (보냈는지 웹메일 보낸메일함도 확인해 주세요)')
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(os.path.join(OUT_DIR, 'sent.log'), 'a', encoding='utf-8') as log:
        log.write('%s\t%s\t%s\n' % (datetime.datetime.now(KST).strftime('%Y-%m-%d %H:%M'), ', '.join(draft['to'] + draft['cc']), draft['subject']))
    print(json.dumps({'sent': True, 'to': draft['to'], 'cc': draft['cc'], 'subject': draft['subject']}, ensure_ascii=False))


# ───────────────────────── 화면(HTML) ─────────────────────────
def esc(value):
    return html.escape(str(value if value is not None else ''), quote=True)


def sender_name(value):
    return re.sub(r'<[^>]*>', '', str(value or '')).replace('"', '').strip() or '(보낸 사람 없음)'


def when_text(value):
    when = parse_when(value)
    return '%d.%d %02d:%02d' % (when.month, when.day, when.hour, when.minute) if when else str(value or '')[:25]


QUOTE = re.compile(r'\n\s*-{3,}\s*Original Message\s*-{3,}', re.I)


def split_body(body):
    body = str(body or '')
    match = QUOTE.search(body)
    return (body[:match.start()].strip(), body[match.start():].strip()) if match else (body.strip(), '')


def deg(value):
    return '-' if value is None else '%d°' % round(value)


def weather_html(w):
    if w is None:
        return ''
    if 'error' in w:
        return '<section class="panel"><h2>🌡️ 오늘 날씨</h2><p class="muted">%s</p></section>' % esc(w['error'])
    extra = ' · '.join(x for x in [
        '체감 ' + deg(w['feels']) if w['feels'] is not None else '',
        '습도 %d%%' % round(w['humidity']) if w['humidity'] is not None else '',
        '바람 %.1fm/s' % w['wind'] if w['wind'] is not None else ''] if x)
    rain = ''
    if w['rain'] is not None:
        rain = ' · 강수 확률 %d%%' % round(w['rain'])
    tip = '<p class="tip">☂️ 우산을 챙기세요.</p>' if (w['rain'] or 0) >= 50 else ''
    return ('<section class="panel"><h2>%s %s 오늘 날씨</h2><p class="big">%s <b>%s</b> · %s</p>'
            '<p>최고 %s / 최저 %s%s</p><p class="muted">%s</p>%s'
            '<p class="muted small">Weather data by Open-Meteo.com</p></section>') % (
        w['icon'], esc(w['name']), w['icon'], esc(deg(w['temp'])), esc(w['label']), esc(deg(w['max'])), esc(deg(w['min'])), esc(rain), esc(extra), tip)


def law_html(law):
    if law is None:
        return ''
    if 'error' in law:
        return '<section class="panel"><h2>⚖ 법령 개정</h2><p class="muted">%s</p></section>' % esc(law['error'])
    if not law['blocks']:
        return '<section class="panel"><h2>⚖ 법령 개정</h2><p>최근 %d일 안에 새로 공포된 개정이 없습니다.</p><p class="muted small">확인 대상: %s</p></section>' % (
            law['days'], esc(', '.join(law['laws'])))
    out = []
    for b in law['blocks']:
        rows = ''.join('<div class="chg"><b>%s</b><p class="old">구: %s</p><p class="new">신: %s</p></div>' % (
            esc(c['article'] or '(조문)'), esc(c['old']), esc(c['new'])) for c in b['changes'])
        more = '<p class="muted">… 외 %d곳</p>' % b['more'] if b['more'] else ''
        badge = '<span class="badge">NEW</span>' if b['fresh'] else ''
        out.append('<details class="law"><summary><b>%s</b> %s %s<span class="muted"> · 공포 %s · 시행 %s · 바뀐 조문 %d곳</span></summary>%s%s</details>' % (
            esc(b['name']), esc(b['kind']), badge, esc(b['promulgated']), esc(b['effective'] or '-'), b['count'], rows, more))
    fresh = sum(1 for b in law['blocks'] if b['fresh'])
    title = '⚖ 법령 개정 %d건%s' % (len(law['blocks']), ' (새로 확인 %d건)' % fresh if fresh else '')
    return ('<details class="panel lawbox" open><summary class="ptitle">%s</summary>%s'
            '<p class="muted small">※ 국가법령정보센터 신구조문 대비표 기준입니다. 정확한 내용은 원문을 확인하세요. 건별로 눌러 펼칩니다.</p></details>' % (title, ''.join(out)))


def file_link(path, embedded_total):
    name = os.path.basename(path)
    size = os.path.getsize(path)
    if size <= EMBED_FILE_LIMIT and embedded_total[0] + size <= EMBED_TOTAL_LIMIT:
        # 페이지 한 장만 보내도 첨부를 누르면 내려받을 수 있게 파일을 페이지 안에 넣습니다.
        with open(path, 'rb') as handle:
            data = base64.b64encode(handle.read()).decode('ascii')
        embedded_total[0] += size
        mime = mimetypes.guess_type(name)[0] or 'application/octet-stream'
        return '<a class="file" download="%s" href="data:%s;base64,%s">%s</a>' % (esc(name), esc(mime), data, esc(name))
    rel = '/'.join(urllib.parse.quote(p) for p in path.replace(os.sep, '/').split('/')[-3:])
    return '<a class="file" href="%s" target="_blank" rel="noopener">%s</a>' % (esc(rel), esc(name))


def mail_html(mail, mail_error):
    if mail_error:
        return '<section class="panel warn"><h2>📬 오늘 받은 메일</h2><p>%s</p></section>' % esc(mail_error)
    messages = mail['messages']
    embedded = [0]
    index = ''.join('<a href="#m%d"><span class="n">%d</span><span class="who">%s</span><span class="subj">%s</span><span class="t">%s</span></a>' % (
        i + 1, i + 1, esc(sender_name(m['from'])), esc(m['subject'][:70]), esc(when_text(m['date']))) for i, m in enumerate(messages))
    cards = []
    for i, m in enumerate(messages):
        main, quoted = split_body(m['body'])
        saved = [p for p in m['savedFiles'] if os.path.exists(p)]
        saved_names = set(os.path.basename(p) for p in saved)
        files = ''.join(file_link(p, embedded) for p in saved)
        others = ''.join('<span class="file off" title="받지 않은 형식이거나 용량 초과">%s</span>' % esc(n) for n in m['attachments']
                         if re.sub(r'[^0-9A-Za-z가-힣._() -]', '_', n)[:80] not in saved_names)
        cards.append('<article class="card" id="m%d"><header><span class="no">%d</span><div><p class="from">%s</p><h3>%s</h3><p class="muted">%s%s</p></div></header>'
                     '<pre class="body">%s</pre>%s%s</article>' % (
                         i + 1, i + 1, esc(sender_name(m['from'])), esc(m['subject'] or '(제목 없음)'), esc(when_text(m['date'])), ' · 참조 있음' if m['cc'] else '',
                         esc(main) or '(본문 없음)',
                         '<details><summary>이전 메일 내용 보기</summary><pre class="body quoted">%s</pre></details>' % esc(quoted) if quoted else '',
                         '<div class="files"><span class="label">첨부</span>%s%s</div>' % (files, others) if files or others else ''))
    limit = '<p class="muted">※ 받은편지함 최근 %d통까지만 확인했습니다. 더 이른 오늘 메일은 웹메일에서 확인하세요.</p>' % MAX_MAILS_SCANNED if mail['limited'] else ''
    head = '<h2>📬 받은 메일 %d통</h2>' % len(messages)
    skipped = ' · 제외 %d통' % mail['skipped'] if mail['skipped'] else ''
    span = ''
    if mail.get('start') and mail.get('end'):
        span = '%s ~ %s 도착분' % (mail['start'].strftime('%m.%d %H:%M'), mail['end'].strftime('%m.%d %H:%M'))
    later = ' · 이후 도착 %d통은 제외' % mail['later'] if mail.get('later') else ''
    sub = '<p class="muted">%s · 읽기만 했고 삭제·발송은 하지 않았습니다%s%s</p>' % (esc(span), skipped, later)
    body = ('<nav class="index">%s</nav>' % index if messages else '<p>이 구간에 받은 메일이 없습니다.</p>') + '\n'.join(cards)
    return '<section>%s%s%s%s</section>' % (head, sub, limit, body)


CSS = """:root{--bg:#f4f5f7;--card:#fff;--text:#1d2129;--sub:#6b7280;--line:#e3e6eb;--accent:#2d5be3;--chip:#eef2ff;--warn:#b45309;--old:#9a3412;--new:#166534}
@media (prefers-color-scheme:dark){:root{--bg:#14161a;--card:#1d2026;--text:#e8eaee;--sub:#9aa3b2;--line:#2c313a;--accent:#8aa4ff;--chip:#262c3d;--warn:#fbbf24;--old:#fdba74;--new:#86efac}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:15px/1.7 -apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif}
main{max-width:860px;margin:0 auto;padding:24px 16px 64px}h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:0 0 8px}h3{font-size:15px;margin:0;line-height:1.5}
.sub{color:var(--sub);margin:0 0 16px;font-size:13px}.muted{color:var(--sub);margin:2px 0}.small{font-size:12px}.big{font-size:20px;margin:4px 0}
.panel{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:14px 18px;margin:0 0 14px}.panel p{margin:4px 0}.warn{border-color:var(--warn)}.tip{color:var(--accent);font-weight:600}
.ptitle{font-size:17px;font-weight:700;cursor:pointer}.lawbox>summary{margin-bottom:6px}.law{padding:8px 0;border-top:1px dashed var(--line)}.law>summary{cursor:pointer;color:var(--text);font-size:14px;list-style-position:inside}.law[open]>summary{margin-bottom:4px}.chg{margin:8px 0 0 8px;font-size:13px}.chg p{margin:2px 0}.old{color:var(--old)}.new{color:var(--new)}
.badge{display:inline-block;margin-left:6px;padding:0 7px;border-radius:999px;background:var(--accent);color:var(--bg);font-size:11px;font-weight:700;vertical-align:middle}
.index{display:grid;gap:2px;margin:10px 0 18px;padding:6px;background:var(--card);border:1px solid var(--line);border-radius:12px}
.index a{display:grid;grid-template-columns:26px minmax(90px,150px) 1fr auto;gap:10px;align-items:baseline;padding:8px 10px;border-radius:8px;color:var(--text);text-decoration:none}.index a:hover{background:var(--chip)}
.index .n{color:var(--accent);font-weight:700;font-size:13px}.index .who{font-weight:700;font-size:15px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.index .subj{color:var(--sub);font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.index .t{color:var(--sub);font-size:12px;white-space:nowrap}
.from{margin:0;font-size:16px;font-weight:700;color:var(--accent)}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px 18px;margin:0 0 14px}
.card header{display:flex;gap:12px;align-items:flex-start}.no{flex:none;width:26px;height:26px;border-radius:50%;background:var(--chip);color:var(--accent);display:grid;place-items:center;font-weight:700;font-size:13px}
.body{white-space:pre-wrap;word-break:break-word;font:inherit;margin:12px 0 0}.quoted{color:var(--sub);font-size:13px}
details{margin-top:10px}summary{cursor:pointer;color:var(--accent);font-size:13px}
.files{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:14px;padding-top:12px;border-top:1px dashed var(--line)}.label{color:var(--sub);font-size:12px}
.file{display:inline-block;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:5px 11px;border-radius:999px;background:var(--chip);color:var(--accent);text-decoration:none;font-size:13px}.file:hover{text-decoration:underline}.file.off{color:var(--sub);background:transparent;border:1px solid var(--line)}"""


def render_page(weather, law, mail, mail_error, notes):
    now = datetime.datetime.now(KST)
    weekday = '월화수목금토일'[now.weekday()]
    short = '%02d.%02d.%02d' % (now.year % 100, now.month, now.day)
    note = ''.join('<p class="panel warn">%s</p>' % esc(n) for n in notes)
    return ('<!doctype html>\n<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            '<title>아침 브리핑_%s</title><style>%s</style></head><body><main><h1>아침 브리핑 <span class="muted">%s</span></h1><p class="sub">%d년 %d월 %d일 (%s) %02d:%02d 기준</p>%s%s%s%s</main></body></html>') % (
        short, CSS, short, now.year, now.month, now.day, weekday, now.hour, now.minute, note, weather_html(weather), law_html(law), mail_html(mail, mail_error))


DATED_PAGE = re.compile(r'아침 브리핑_(\d{2})\.(\d{2})\.(\d{2})\.html')


def clean_old_pages(out_dir, keep_days, today=None):
    """날짜가 붙은 오래된 브리핑 파일만 지웁니다. 이 스크립트가 만든 '아침 브리핑_YY.MM.DD.html' 만 대상이고,
    날짜는 파일 이름에서 읽습니다. 다른 파일, 폴더, 바로가기(심볼릭 링크)는 건드리지 않습니다."""
    if not isinstance(keep_days, int) or isinstance(keep_days, bool) or not 1 <= keep_days <= 365:
        keep_days = 7
    today = today or datetime.datetime.now(KST).date()
    removed = 0
    try:
        names = os.listdir(out_dir)
    except OSError:
        return 0
    for name in names:
        m = DATED_PAGE.fullmatch(name)
        path = os.path.join(out_dir, name)
        if not m or os.path.islink(path) or not os.path.isfile(path):
            continue
        try:
            made = datetime.date(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            continue
        if (today - made).days > keep_days:
            try:
                os.remove(path)
                removed += 1
            except OSError:
                pass
    return removed


# ───────────────────────── 실행 ─────────────────────────
def demo_data():
    weather = {'name': '서울', 'label': '구름 조금', 'icon': '⛅', 'temp': 12.0, 'feels': 10.0, 'humidity': 55.0, 'wind': 2.3, 'max': 16.0, 'min': 8.0, 'rain': 60.0}
    law = {'days': 14, 'laws': DEFAULT_LAWS, 'blocks': [{'name': '산업안전보건기준에 관한 규칙', 'kind': '일부개정', 'promulgated': '2026-10-20', 'effective': '2027-01-01',
                                                         'count': 1, 'more': 0, 'fresh': True,
                                                         'changes': [{'article': '제00조(예시)', 'old': '(예시) 종전 문장입니다.', 'new': '(예시) 개정된 문장입니다.'}]}]}
    mail = {'skipped': 1, 'limited': False, 'messages': [
        {'from': '"홍길동" <hong@example.com>', 'to': 'me@example.com', 'cc': '', 'subject': '[예시] 회의 일정 확인 부탁드립니다', 'date': 'Thu, 29 Oct 2026 09:20:00 +0900',
         'body': '안녕하세요.\n내일 오후 2시 회의가 가능한지 확인 부탁드립니다.\n\n--- Original Message ---\n이전 메일 내용(예시)', 'attachments': ['예시자료.pdf'], 'savedFiles': []}]}
    return weather, law, mail


def main(argv):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8')  # 윈도우 콘솔에서 한글이 깨지거나 오류가 나지 않게
        except Exception:
            pass
    if '--init' in argv:
        return init_files()
    if '--demo' in argv:
        os.makedirs(OUT_DIR, exist_ok=True)
        weather, law, mail = demo_data()
        write_private(os.path.join(OUT_DIR, 'mail-today.html'), render_page(weather, law, mail, '', ['(예시 화면입니다. 실제 메일이 아닙니다)']))
        print(json.dumps({'page': 'out/mail-today.html', 'demo': True}, ensure_ascii=False))
        return
    if '--send' in argv:
        pos = argv.index('--send')
        if pos + 1 >= len(argv):
            raise SystemExit('사용법: python3 mail-today.py --send 초안.json [--confirm 코드]')
        confirm = argv[argv.index('--confirm') + 1] if '--confirm' in argv and argv.index('--confirm') + 1 < len(argv) else None
        return send_flow(argv[pos + 1], confirm, read_json(SETTINGS_FILE, {}) or {}, read_json(STATE_FILE, {}) or {})
    account, password = load_account()
    if '--check' in argv:
        try:
            check_login(account, password)
        except Exception:
            raise SystemExit('하이웍스 로그인에 실패했습니다. 아이디·비밀번호와 POP3 사용 설정(웹메일 환경설정)을 확인해 주세요.')
        print(json.dumps({'connected': True}, ensure_ascii=False))
        return
    settings = read_json(SETTINGS_FILE, {}) or {}
    state = read_json(STATE_FILE, {}) or {}
    notes = [x for x in [permission_warning()] if x]
    attach_dir = os.path.join(OUT_DIR, 'mail-today-files')
    # 첨부는 실행할 때마다 지우고 새로 받습니다. (이 폴더 밖은 건드리지 않습니다)
    if os.path.isdir(attach_dir):
        for root, dirs, files in os.walk(attach_dir, topdown=False):
            for f in files:
                os.remove(os.path.join(root, f))
            for d in dirs:
                os.rmdir(os.path.join(root, d))
    os.makedirs(attach_dir, exist_ok=True)
    mail, mail_error = None, ''
    try:
        mail = read_today(account, password, as_list(settings.get('exclude')), attach_dir, mail_window(settings))
    except Exception:
        mail_error = '메일을 읽지 못했습니다. 아이디·비밀번호, 하이웍스의 POP3 사용 설정, 인터넷 연결을 확인해 주세요.'
    weather = get_weather(settings.get('weather', {}))
    law = get_law(settings, state)
    page = render_page(weather, law, mail, mail_error, notes)
    write_private(os.path.join(OUT_DIR, 'mail-today.html'), page)
    write_private(os.path.join(OUT_DIR, '아침 브리핑_%s.html' % datetime.datetime.now(KST).strftime('%y.%m.%d')), page)
    cleaned = clean_old_pages(OUT_DIR, settings.get('keepDays', 7))
    law_brief = None
    if law is not None:
        if 'error' in law:
            law_brief = {'error': law['error']}
        else:
            law_brief = {'days': law['days'], 'count': len(law['blocks']),
                         'items': ['%s %s (공포 %s, 시행 %s, 바뀐 조문 %d곳)' % (b['name'], b['kind'], b['promulgated'], b['effective'] or '-', b['count']) for b in law['blocks']]}
    weather_brief = None
    if weather is not None:
        if 'error' in weather:
            weather_brief = {'error': weather['error']}
        else:
            weather_brief = {'place': weather['name'], 'now': weather['temp'], 'max': weather['max'], 'min': weather['min'],
                             'label': weather['label'], 'rainPercent': weather['rain'], 'feels': weather['feels']}
    # 답장 초안을 쓸 때 참고하는 자료입니다. (화면의 번호와 같은 순서, 이 컴퓨터에만 저장)
    write_private(os.path.join(OUT_DIR, 'mail-today.json'), json.dumps(
        {'messages': mail['messages'] if mail else [], 'weather': weather_brief, 'law': law_brief}, ensure_ascii=False, indent=1))
    write_private(STATE_FILE, json.dumps(state, ensure_ascii=False))
    print(json.dumps({'page': 'out/mail-today.html', 'mails': len(mail['messages']) if mail else 0, 'skipped': mail['skipped'] if mail else 0,
                      'mailError': bool(mail_error), 'cleaned': cleaned, 'weather': weather_brief, 'law': law_brief}, ensure_ascii=False))
    if mail_error:
        sys.exit(1)


if __name__ == '__main__':
    main(sys.argv[1:])
