from __future__ import annotations

import hashlib
import hmac
import json
from urllib.parse import parse_qs

def validate_init_data(init_data: str, bot_token: str) -> dict | None:
    try:
        vals = {k: v[0] for k, v in parse_qs(init_data).items()}
        hash_val = vals.pop('hash', None)
        if not hash_val:
            return None

        data_check_string = '\n'.join(f'{k}={v}' for k, v in sorted(vals.items()))
        secret_key = hmac.new(b'WebAppData', bot_token.encode(), hashlib.sha256).digest()
        calculated_hash = hmac.new(secret_key, data_check_string.encode(), hashlib.sha256).hexdigest()

        if calculated_hash != hash_val:
            return None

        user_data = vals.get('user')
        if user_data:
            return json.loads(user_data)
        return None
    except Exception:
        return None
