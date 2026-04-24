"""Tests for _openai_chat_with_retry — exponential backoff wrapper."""
import pytest


@pytest.fixture
def fn():
    import app
    return app._openai_chat_with_retry


class FakeClient:
    """Raises `fail_times` transient errors then returns `return_value`."""
    def __init__(self, error_factory, fail_times=0, return_value="ok"):
        self.fail_times = fail_times
        self.calls = 0
        self.return_value = return_value
        self.error_factory = error_factory
        self.chat = self
        self.completions = self

    def create(self, **kwargs):
        self.calls += 1
        if self.calls <= self.fail_times:
            raise self.error_factory()
        return self.return_value


def _rate_limit_err():
    import httpx
    from openai import RateLimitError
    req = httpx.Request("POST", "http://test")
    resp = httpx.Response(429, request=req, content=b'{"error":"rate"}')
    return RateLimitError(message="rate limit", response=resp, body={"error": "rate"})


def _timeout_err():
    import httpx
    from openai import APITimeoutError
    req = httpx.Request("POST", "http://test")
    return APITimeoutError(request=req)


def test_succeeds_on_first_call(fn, monkeypatch):
    monkeypatch.setattr("app.OPENAI_MAX_RETRIES", 3)
    monkeypatch.setattr("app.OPENAI_RETRY_INITIAL_WAIT", 0.0)
    client = FakeClient(_rate_limit_err, fail_times=0, return_value="done")
    assert fn(client, foo="bar") == "done"
    assert client.calls == 1


def test_retries_then_succeeds(fn, monkeypatch):
    monkeypatch.setattr("app.OPENAI_MAX_RETRIES", 3)
    monkeypatch.setattr("app.OPENAI_RETRY_INITIAL_WAIT", 0.0)
    monkeypatch.setattr("app.OPENAI_RETRY_MAX_WAIT", 0.0)
    client = FakeClient(_rate_limit_err, fail_times=2, return_value="ok")
    assert fn(client) == "ok"
    assert client.calls == 3  # 2 failures + 1 success


def test_gives_up_after_max_retries(fn, monkeypatch):
    monkeypatch.setattr("app.OPENAI_MAX_RETRIES", 2)
    monkeypatch.setattr("app.OPENAI_RETRY_INITIAL_WAIT", 0.0)
    client = FakeClient(_rate_limit_err, fail_times=99)
    from openai import RateLimitError
    with pytest.raises(RateLimitError):
        fn(client)
    assert client.calls == 3  # initial + 2 retries


def test_retries_on_timeout(fn, monkeypatch):
    monkeypatch.setattr("app.OPENAI_MAX_RETRIES", 3)
    monkeypatch.setattr("app.OPENAI_RETRY_INITIAL_WAIT", 0.0)
    client = FakeClient(_timeout_err, fail_times=1, return_value="ok")
    assert fn(client) == "ok"
    assert client.calls == 2


def test_non_retryable_error_propagates_immediately(fn, monkeypatch):
    monkeypatch.setattr("app.OPENAI_MAX_RETRIES", 3)
    monkeypatch.setattr("app.OPENAI_RETRY_INITIAL_WAIT", 0.0)

    def _bad_input():
        return ValueError("bad prompt")

    client = FakeClient(_bad_input, fail_times=5)
    with pytest.raises(ValueError):
        fn(client)
    assert client.calls == 1  # no retries on ValueError


def test_exponential_backoff_waits_grow(fn, monkeypatch):
    """Verify the wait doubles on each retry (capped at OPENAI_RETRY_MAX_WAIT)."""
    monkeypatch.setattr("app.OPENAI_MAX_RETRIES", 3)
    monkeypatch.setattr("app.OPENAI_RETRY_INITIAL_WAIT", 1.0)
    monkeypatch.setattr("app.OPENAI_RETRY_MAX_WAIT", 10.0)

    sleep_calls = []
    import time as _t
    monkeypatch.setattr(_t, "sleep", lambda s: sleep_calls.append(s))

    client = FakeClient(_rate_limit_err, fail_times=3, return_value="ok")
    assert fn(client) == "ok"
    # 3 retries → 3 sleeps: 1, 2, 4
    assert sleep_calls == [1.0, 2.0, 4.0]
