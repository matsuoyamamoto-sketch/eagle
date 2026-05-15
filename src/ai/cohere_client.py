"""Cohere API クライアント (レート制御 + リトライ + JSON 検証)。"""
from __future__ import annotations

import json
import threading
import time
from collections import deque
from typing import Any, Callable

import cohere

from ..config import settings


class RateLimiter:
    """直近 60 秒で N 回まで許可するスロットリング。"""

    def __init__(self, requests_per_minute: int) -> None:
        self.rpm = max(1, requests_per_minute)
        self._times: deque[float] = deque()
        self._lock = threading.Lock()

    def acquire(self, on_wait: Callable[[float], None] | None = None) -> None:
        with self._lock:
            now = time.monotonic()
            while self._times and now - self._times[0] > 60.0:
                self._times.popleft()
            if len(self._times) >= self.rpm:
                wait = 60.0 - (now - self._times[0]) + 0.05
                wait = max(wait, 0)
                if on_wait and wait > 0.1:
                    on_wait(wait)
                time.sleep(wait)
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
        self.event_hook: Callable[[dict], None] | None = None

    def _emit(self, **kwargs: Any) -> None:
        if self.event_hook:
            try:
                self.event_hook(kwargs)
            except Exception:
                pass

    def _do_chat(self, system: str, user: str, schema: dict[str, Any] | None) -> str:
        self._limiter.acquire(on_wait=lambda w: self._emit(phase="rate_wait", wait=w))
        self._emit(phase="request_start")
        t0 = time.monotonic()
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
        # ハートビート: 別スレッドで一定間隔ごとに経過秒を通知
        stop_hb = threading.Event()

        def _heartbeat() -> None:
            interval = 15.0  # 秒
            next_at = interval
            while not stop_hb.wait(1.0):
                el = time.monotonic() - t0
                if el >= next_at:
                    note = ""
                    if el >= 60:
                        note = " (コールドスタートの可能性あり)"
                    elif el >= 30:
                        note = " (応答待ち継続中…)"
                    self._emit(phase="heartbeat", elapsed=el, note=note)
                    next_at += interval

        hb_thread = threading.Thread(target=_heartbeat, daemon=True)
        hb_thread.start()

        try:
            resp = self._client.chat(**kwargs)
        except Exception as e:
            stop_hb.set()
            self._emit(phase="request_error", elapsed=time.monotonic() - t0, error=str(e))
            raise
        finally:
            stop_hb.set()
        elapsed = time.monotonic() - t0
        try:
            text = resp.message.content[0].text  # type: ignore[attr-defined]
        except Exception:
            text = str(resp)
        self._emit(phase="request_end", elapsed=elapsed, chars=len(text))
        return text

    def chat_json(
        self,
        system: str,
        user: str,
        schema: dict[str, Any] | None = None,
    ) -> Any:
        max_retries = max(1, settings.cohere_max_retries)
        last_err: Exception | None = None
        for attempt in range(1, max_retries + 1):
            try:
                raw = self._do_chat(system, user, schema)
                try:
                    return json.loads(raw)
                except json.JSONDecodeError as e:
                    raise RuntimeError(
                        f"Cohere 応答の JSON パースに失敗: {e}\n--- raw ---\n{raw[:500]}"
                    )
            except Exception as e:
                last_err = e
                if attempt >= max_retries:
                    raise
                wait = min(30.0, 2.0 * (2 ** (attempt - 1)))
                self._emit(phase="retry", attempt=attempt, max=max_retries,
                           wait=wait, error=str(e)[:200])
                time.sleep(wait)
        if last_err:
            raise last_err
        raise RuntimeError("unreachable")
