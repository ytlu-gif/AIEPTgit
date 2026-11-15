"""Read-only checks of this project's public, stable GitHub releases."""
import re
from urllib.parse import quote

import requests

LATEST_URL = 'https://api.github.com/repos/ytlu-gif/AIEPTgit/releases/latest'
RELEASES_URL = 'https://github.com/ytlu-gif/AIEPTgit/releases'


def version_tuple(value):
    if not isinstance(value, str):
        return None
    match = re.fullmatch(r'v?(\d+)\.(\d+)(?:\.(\d+))?', value.strip())
    return tuple(int(x or 0) for x in match.groups()) if match else None


def should_check(enabled, last_attempt, now):
    if not enabled:
        return False
    if not isinstance(last_attempt, (int, float)) or last_attempt <= 0:
        return True
    elapsed = now - last_attempt
    return not 0 <= elapsed < 86400


def check_latest(current_version):
    current = version_tuple(current_version)
    if current is None:
        raise ValueError('Invalid installed version')
    try:
        with requests.Session() as session:
            # No API keys, lessons, transcripts, or user identifiers are sent.
            session.trust_env = False
            response = session.get(LATEST_URL,
                headers={'Accept': 'application/vnd.github+json', 'User-Agent': 'AIEPT-update-check'},
                timeout=(3, 5), allow_redirects=False)
        if response.status_code == 404:
            return {'status': 'no_release'}
        if response.status_code in (403, 429):
            raise RuntimeError('GitHub 暫時限制版本查詢，請稍後再試。')
        if response.status_code != 200:
            raise RuntimeError(f'版本查詢暫時無法使用（HTTP {response.status_code}）。')
        data = response.json()
        if not isinstance(data, dict):
            raise ValueError('Invalid release response')
        if data.get('draft') or data.get('prerelease'):
            return {'status': 'no_release'}
        tag = data.get('tag_name')
        latest = version_tuple(tag)
        if latest is None:
            raise RuntimeError('無法辨識 Release 的版本標籤，請至 GitHub 發行頁查看。')
        return {'status': 'available' if latest > current else 'current', 'tag': tag,
                'url': RELEASES_URL + '/tag/' + quote(tag.strip(), safe='')}
    except requests.RequestException as error:
        raise RuntimeError('無法連線 GitHub 檢查更新；請確認網路，或稍後手動檢查。') from error
    except (ValueError, TypeError) as error:
        raise RuntimeError('GitHub 版本資料格式不正確，請稍後再試。') from error
