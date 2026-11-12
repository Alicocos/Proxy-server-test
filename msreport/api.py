"""Минимальный клиент JSON API МойСклад 1.2 (только чтение)."""
import time

import requests

BASE = "https://api.moysklad.ru/api/remap/1.2"


class ApiError(RuntimeError):
    pass


class MoySklad:
    def __init__(self, token: str):
        self.s = requests.Session()
        self.s.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept-Encoding": "gzip",
            "Accept": "application/json;charset=utf-8",
        })

    def get(self, path: str, params: dict | None = None) -> dict:
        url = path if path.startswith("http") else f"{BASE}/{path.lstrip('/')}"
        for attempt in range(6):
            try:
                r = self.s.get(url, params=params, timeout=120)
            except requests.RequestException as e:
                if attempt == 5:
                    raise ApiError(f"Нет связи с МойСклад: {e}") from e
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 429 or r.status_code >= 500:
                wait_ms = r.headers.get("X-Lognex-Retry-After") or r.headers.get("X-Lognex-Retry-TimeInterval")
                time.sleep(max(float(wait_ms) / 1000 if wait_ms else 0, 1.5 * (attempt + 1)))
                continue
            if r.status_code == 401:
                raise ApiError("МойСклад отклонил токен (401). Проверьте MOYSKLAD_TOKEN.")
            if r.status_code >= 400:
                raise ApiError(f"Ошибка МойСклад {r.status_code} на {path}: {r.text[:500]}")
            return r.json()
        raise ApiError(f"МойСклад не отвечает на {path} после нескольких попыток")

    def rows(self, path: str, params: dict | None = None, limit: int = 1000) -> list[dict]:
        """Все строки с постраничной выгрузкой. При expand лимит МойСклад — 100."""
        params = dict(params or {})
        if "expand" in params:
            limit = min(limit, 100)
        out, offset = [], 0
        while True:
            page = self.get(path, {**params, "limit": limit, "offset": offset})
            chunk = page.get("rows", [])
            out.extend(chunk)
            offset += limit
            size = page.get("meta", {}).get("size")
            if len(chunk) < limit or (size is not None and offset >= size):
                return out
