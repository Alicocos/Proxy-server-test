"""Минимальный клиент JSON API МойСклад 1.2 (только чтение)."""
import time

import requests

BASE = "https://api.moysklad.ru/api/remap/1.2"


class ApiError(RuntimeError):
    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


class MoySklad:
    def __init__(self, token: str, base: str = BASE):
        self.base = base.rstrip("/")
        self.s = requests.Session()
        self.s.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept-Encoding": "gzip",
            "Accept": "application/json;charset=utf-8",
        })

    def get(self, path: str, params: dict | None = None) -> dict:
        url = path if path.startswith("http") else f"{self.base}/{path.lstrip('/')}"
        for attempt in range(10):
            try:
                r = self.s.get(url, params=params, timeout=120)
            except requests.RequestException as e:
                if attempt == 9:
                    raise ApiError(f"Нет связи с МойСклад: {e}") from e
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 429 or r.status_code >= 500:
                wait_ms = r.headers.get("X-Lognex-Retry-After") or r.headers.get("X-Lognex-Retry-TimeInterval")
                time.sleep(max(float(wait_ms) / 1000 if wait_ms else 0, min(3 * (attempt + 1), 20)))
                continue
            if r.status_code == 401:
                raise ApiError("МойСклад отклонил токен (401). Проверьте MOYSKLAD_TOKEN.")
            if r.status_code >= 400:
                raise ApiError(f"Ошибка МойСклад {r.status_code} на {path}: {r.text[:500]}", r.status_code)
            return r.json()
        raise ApiError(f"МойСклад не отвечает на {path} после нескольких попыток")

    def rows(self, path: str, params: dict | None = None, limit: int = 1000, workers: int = 4) -> list[dict]:
        """Все строки с постраничной выгрузкой: первая страница даёт общее число, остальные качаются параллельно
        (МойСклад разрешает до 5 одновременных запросов). При expand лимит страницы — 100."""
        from concurrent.futures import ThreadPoolExecutor
        params = dict(params or {})
        if "expand" in params:
            limit = min(limit, 100)
        first = self.get(path, {**params, "limit": limit, "offset": 0})
        out = list(first.get("rows", []))
        size = first.get("meta", {}).get("size")
        if size is None:  # отчёты без общего числа строк — по очереди
            offset = limit
            chunk = out
            while len(chunk) == limit:
                chunk = self.get(path, {**params, "limit": limit, "offset": offset}).get("rows", [])
                out.extend(chunk)
                offset += limit
            return out
        offsets = list(range(limit, size, limit))
        if offsets:
            with ThreadPoolExecutor(max_workers=workers) as pool:
                for page in pool.map(lambda o: self.get(path, {**params, "limit": limit, "offset": o}), offsets):
                    out.extend(page.get("rows", []))
        return out
