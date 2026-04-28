"""Cohere API クライアント (レート制御 + リトライ + JSON 検証)。"""
from __future__ import annotations

import json
import threading
import time
from collections import deque
from typing import Any

import cohere
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

from ..config import settings


class RateLimiter:
    """直近 60 秒で N 回まで許可するスロットリング。"""

    def __init__(self, requests_per_minute: int) -> None:
        self.rpm = max(1, requests_per_minute)
        self._times: deque[float] = deque()
        self._lock = threading.Lock()

    def acquire(self) -> None:
        with self._lock:
            now = time.monotonic()
            while self._times and now - self._times[0] > 60.0:
                self._times.popleft()
            if len(self._times) >= self.rpm:
                wait = 60.0 - (now - self._times[0]) + 0.05
                time.sleep(max(wait, 0))
                now = time.monotonic()
                while self._times and now - self._times[0] > 60.0:
                    self._times.popleft()
            self._times.append(time.monotonic())


class CohereJSONClient:
    """JSON 応答に特化した Cohere ラッパ。"""

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        rpm: int | None = None,
    ) -> None:
        self.api_key = api_key or settings.cohere_api_key
        if not self.api_key:
            raise RuntimeError("COHERE_API_KEY が設定されていません (.env を確認)")
        self.model = model or settings.cohere_model
        self.rpm = rpm or settings.cohere_requests_per_minute
        self._client = cohere.ClientV2(api_key=self.api_key)
        self._limiter = RateLimiter(self.rpm)

    def _do_chat(self, system: str, user: str, schema: dict[str, Any] | None) -> str:
        self._limiter.acquire()
        kwargs: dict[str, Any] = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "temperature": 0.2,
        }
        if schema is not None:
            kwargs["response_format"] = {"type": "json_object", "schema": schema}
        else:
            kwargs["response_format"] = {"type": "json_object"}
        resp = self._client.chat(**kwargs)
        # ClientV2 の応答形式
        try:
            return resp.message.content[0].text  # type: ignore[attr-defined]
        except Exception:
            return str(resp)

    @retry(
        retry=retry_if_exception_type(Exception),
        stop=stop_after_attempt(settings.cohere_max_retries),
        wait=wait_exponential(multiplier=2, min=2, max=30),
        reraise=True,
    )
    def chat_json(
        self,
        system: str,
        user: str,
        schema: dict[str, Any] | None = None,
    ) -> Any:
        raw = self._do_chat(system, user, schema)
        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Cohere 応答の JSON パースに失敗: {e}\n--- raw ---\n{raw[:500]}")
