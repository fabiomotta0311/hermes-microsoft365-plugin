"""The sync/async seam under the host's default in-process plugin execution.

Hermes invokes plugin tool handlers on a thread that already has a running asyncio
event loop under ``plugins.isolation: in_process`` -- the default. The seam used to
treat that as misuse and fail closed, so every live call died with
``internal_error: cannot block on an async graph call inside a running event loop``
before any Microsoft traffic (issue #16).

These tests run the seam the way the host does: from a thread that owns a running
loop. They also pin the properties the fix must not give up -- the taxonomy still
sanitizes, the caller's contextvars still resolve on the worker, and the seam still
never deadlocks.
"""
from __future__ import annotations

import asyncio
import contextvars
import threading

import pytest


def test_the_seam_completes_a_call_from_a_thread_that_owns_a_running_loop():
    """The live deployment: a running loop on this thread, and the call still works.

    Before the fix this raised ``internal_error``; the coroutine never ran.
    """
    from microsoft365.execution import run_async

    async def answer():
        await asyncio.sleep(0)
        return 42

    result = {}

    def caller():
        async def main():
            result["value"] = run_async(answer())

        asyncio.run(main())

    thread = threading.Thread(target=caller)
    thread.start()
    thread.join(timeout=30)

    assert not thread.is_alive(), "the seam deadlocked the calling thread"
    assert result["value"] == 42


def test_the_seam_completes_a_call_from_inside_a_coroutine_that_owns_the_loop():
    """The exact shape the host produces: a coroutine calling the seam synchronously.

    ``asyncio.run(main())`` puts the caller inside a running loop, and the handler
    calls ``run_async`` from there -- the call path that failed in every live report.
    """
    from microsoft365.execution import run_async

    async def answer():
        return "succeeded"

    async def main():
        return run_async(answer())

    assert asyncio.run(main()) == "succeeded"


def test_a_failure_on_the_worker_path_is_still_sanitized_by_the_taxonomy():
    """The worker path must not leak raw text -- it takes the same exit as the direct one.

    The direct path converts an SDK failure into a typed error whose message is a
    static template. The worker path is a different execution route to the same seam,
    so it must not become a way around the leakage barrier.
    """
    from microsoft365.execution import ExecutionError, run_async

    async def boom():
        raise ValueError("client_secret=SENTINEL-raw-sdk-text")

    async def main():
        return run_async(boom(), category="transport_error")

    with pytest.raises(ExecutionError) as caught:
        asyncio.run(main())

    assert caught.value.category == "transport_error"
    assert "SENTINEL-raw-sdk-text" not in str(caught.value)
    assert isinstance(caught.value.__cause__, ValueError)


def test_the_caller_contextvars_reach_the_worker_thread():
    """The worker's contextvars are copied from the caller's.

    The host resolves the profile and the approval/sudo callbacks from contextvars on
    the calling thread. A worker that started with a blank context would resolve a
    different profile, so the copy is what keeps the worker's call identical to the
    caller's.
    """
    from microsoft365.execution import run_async

    marker = contextvars.ContextVar("marker", default="unset")

    async def read_marker():
        return marker.get()

    async def main():
        marker.set("carried")
        return run_async(read_marker())

    assert asyncio.run(main()) == "carried"


def test_the_worker_thread_does_not_leak_its_event_loop():
    """Every worker call closes the loop it created.

    The seam's guarantee is that no loop, client or credential is shared between
    calls. A worker that left its loop open would break that for the next call, and
    would leak a running loop per call on a long-lived host thread.
    """
    from microsoft365.execution import run_async

    seen = []

    async def answer():
        seen.append(asyncio.get_running_loop())
        return 1

    async def main():
        return run_async(answer())

    asyncio.run(main())

    assert len(seen) == 1
    assert seen[0].is_closed(), "the worker left its event loop open"


def test_each_worker_call_gets_a_fresh_loop():
    """Two calls do not share a loop, so state cannot leak between them."""
    from microsoft365.execution import run_async

    seen = []

    async def answer():
        seen.append(asyncio.get_running_loop())
        return 1

    async def main():
        run_async(answer())
        run_async(answer())

    asyncio.run(main())

    assert len(seen) == 2
    assert seen[0] is not seen[1]


def test_the_seam_still_rejects_a_non_awaitable():
    """The fail-closed guard on the input side is untouched by the worker path."""
    from microsoft365.execution import ExecutionError, run_async

    with pytest.raises(ExecutionError) as caught:
        run_async(object())

    assert caught.value.category == "configuration_error"


def test_the_seam_does_not_deadlock_when_called_from_its_own_worker():
    """A call made from inside a coroutine the seam is running does not hang.

    The direct path runs the coroutine on the caller's loop; the worker path runs it
    on the worker's loop. Neither may wait for a loop that is itself waiting.
    """
    from microsoft365.execution import run_async

    async def inner():
        return "inner"

    async def outer():
        return run_async(inner())

    async def main():
        return await asyncio.wait_for(outer(), timeout=20)

    assert asyncio.run(main()) == "inner"


def test_a_stuck_call_on_the_worker_is_bounded_and_reported_as_a_transport_error(monkeypatch):
    """A call that never finishes is bounded, not left hanging forever.

    Without the bound, one stuck Graph call would hold the caller's thread forever --
    and under the host's in-process execution that thread is the host's own.
    """
    from microsoft365 import execution
    from microsoft365.execution import ExecutionError, run_async

    monkeypatch.setattr(execution, "_WORKER_TIMEOUT_SECONDS", 1)

    async def stuck():
        await asyncio.sleep(3600)

    async def main():
        return run_async(stuck())

    with pytest.raises(ExecutionError) as caught:
        asyncio.run(main())

    assert caught.value.category == "transport_error"
    assert "SENTINEL" not in str(caught.value)


def test_a_failure_on_the_worker_keeps_the_raw_text_out_of_the_message(monkeypatch):
    """The worker's own failure is sanitized exactly like the direct path's.

    The worker is a different route to the same seam, so it must not become a way
    around the leakage barrier: the category comes from the taxonomy and the raw text
    stays in ``__cause__``.
    """
    from microsoft365.execution import ExecutionError, run_async

    async def boom():
        raise ValueError("client_secret=SENTINEL-raw-sdk-text")

    async def main():
        return run_async(boom(), category="transport_error")

    with pytest.raises(ExecutionError) as caught:
        asyncio.run(main())

    assert caught.value.category == "transport_error"
    assert "SENTINEL-raw-sdk-text" not in str(caught.value)
    assert isinstance(caught.value.__cause__, ValueError)
