import hmac
import hashlib
from urllib.parse import parse_qsl

class TelegramAuthError(Exception):
    pass

def validate_init_data(init_data: str, bot_token: str) -> dict:
    if not init_data:
        raise TelegramAuthError("Empty init_data")

    data = dict(parse_qsl(init_data, keep_blank_values=True))
    received_hash = data.pop("hash", None)
    if not received_hash:
        raise TelegramAuthError("No hash in init_data")

    pairs = [f"{k}={data[k]}" for k in sorted(data.keys())]
    data_check_string = "\n".join(pairs).encode("utf-8")

    secret_key = hmac.new(
        key=b"WebAppData",
        msg=bot_token.encode("utf-8"),
        digestmod=hashlib.sha256
    ).digest()

    calculated_hash = hmac.new(
        key=secret_key,
        msg=data_check_string,
        digestmod=hashlib.sha256
    ).hexdigest()

    if not hmac.compare_digest(calculated_hash, received_hash):
        raise TelegramAuthError("Invalid hash")

    return data
