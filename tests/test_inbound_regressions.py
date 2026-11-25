import json

import jwt
import pytest
from cryptography.hazmat.primitives.asymmetric import rsa

from microsoft365.teams_inbound import InboundTeamsError, INVALID_TOKEN, verify_activity_token


def test_real_signed_token_missing_exp_is_rejected():
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    jwk = json.loads(jwt.algorithms.RSAAlgorithm.to_jwk(key.public_key()))
    jwk['kid'] = 'test-key'
    token = jwt.encode({'aud': 'bot', 'iss': 'https://api.botframework.com'}, key,
                       algorithm='RS256', headers={'kid': 'test-key'})
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(token, audience='bot', keys=[jwk], now=1700000000)
    assert caught.value.status == INVALID_TOKEN


@pytest.mark.parametrize('exp', [None, True, '1700000300', [], float('inf'), float('nan')])
def test_real_signed_token_invalid_exp_is_rejected(exp):
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    jwk = json.loads(jwt.algorithms.RSAAlgorithm.to_jwk(key.public_key()))
    jwk['kid'] = 'test-key'
    token = jwt.encode({'aud': 'bot', 'iss': 'https://api.botframework.com', 'exp': exp}, key,
                       algorithm='RS256', headers={'kid': 'test-key'})
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(token, audience='bot', keys=[jwk], now=1700000000)
    assert caught.value.status == INVALID_TOKEN


def test_failed_handoff_is_retried_then_deduplicated():
    from tests.test_teams_webhook import application, call

    attempts = []
    def handoff(decision):
        attempts.append(decision)
        if len(attempts) == 1:
            raise RuntimeError('temporary failure')

    app = application(on_activity=handoff)
    assert call(app)[0] == 500
    assert call(app)[2] == {'status': 'accepted'}
    assert call(app)[2] == {'status': 'duplicate'}
    assert len(attempts) == 2


def test_concurrent_duplicate_waits_for_handoff_before_acknowledging():
    import threading
    from concurrent.futures import ThreadPoolExecutor
    from tests.test_teams_webhook import application, call

    entered, release = threading.Event(), threading.Event()
    attempts = []
    def handoff(decision):
        attempts.append(decision)
        entered.set()
        assert release.wait(5)

    app = application(on_activity=handoff)
    with ThreadPoolExecutor(max_workers=2) as pool:
        first = pool.submit(call, app)
        assert entered.wait(5)
        second = pool.submit(call, app)
        try:
            assert not second.done()
        finally:
            release.set()
        results = [first.result(timeout=5), second.result(timeout=5)]
    assert sorted(result[2]['status'] for result in results) == ['accepted', 'duplicate']
    assert len(attempts) == 1


def test_concurrent_waiter_retries_failed_handoff():
    import threading
    from concurrent.futures import ThreadPoolExecutor
    from tests.test_teams_webhook import application, call

    entered, release = threading.Event(), threading.Event()
    attempts = []
    def handoff(decision):
        attempts.append(decision)
        if len(attempts) == 1:
            entered.set()
            assert release.wait(5)
            raise InboundTeamsError('validation_error', status='invalid_activity')

    app = application(on_activity=handoff)
    with ThreadPoolExecutor(max_workers=2) as pool:
        first = pool.submit(call, app)
        assert entered.wait(5)
        second = pool.submit(call, app)
        release.set()
        assert first.result(timeout=5)[0] == 500
        assert second.result(timeout=5)[2] == {'status': 'accepted'}
    assert len(attempts) == 2
