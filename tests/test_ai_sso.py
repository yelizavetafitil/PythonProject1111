import hashlib
import hmac
import os
import sys
import unittest
from urllib.parse import parse_qs, urlencode, urlsplit
from unittest.mock import MagicMock, patch

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import app as portal


class PortalAiSsoTests(unittest.TestCase):
    def setUp(self):
        portal.app.config.update(TESTING=True)
        self.client = portal.app.test_client()
        self.settings = patch.multiple(
            portal,
            AI_ASSISTANT_PUBLIC_URL='https://ai.energoprom.by',
            AI_ASSISTANT_CALLBACK_URL='https://ai.energoprom.by/sso-login',
            AI_SSO_SHARED_SECRET='shared-test-secret',
            AI_SSO_MAX_AGE_SEC=300,
        )
        self.settings.start()

    def tearDown(self):
        self.settings.stop()

    def _return_url(self, state='A' * 32):
        return f'https://ai.energoprom.by/sso-login?{urlencode({"state": state})}'

    def test_authenticated_portal_session_returns_signed_ai_callback(self):
        with self.client.session_transaction() as session:
            session['logged_in'] = True
            session['username'] = 'DOMAIN\\ivanov'
            session['display_name'] = 'Иванов Иван'

        response = self.client.get('/login', query_string={'return_url': self._return_url()})
        self.assertEqual(response.status_code, 302)
        location = response.headers['Location']
        parsed = urlsplit(location)
        query = parse_qs(parsed.query)
        self.assertEqual(parsed.netloc, 'ai.energoprom.by')
        self.assertEqual(query['state'][0], 'A' * 32)
        self.assertEqual(query['u'][0], 'ivanov')
        timestamp = query['ts'][0]
        payload = f"ivanov|Иванов Иван|{timestamp}"
        expected = hmac.new(
            b'shared-test-secret', payload.encode('utf-8'), hashlib.sha256
        ).hexdigest()
        self.assertEqual(query['sig'][0], expected)

    def test_login_rejects_external_return_url(self):
        response = self.client.get(
            '/login', query_string={'return_url': 'https://evil.example/sso-login?state=' + 'A' * 32}
        )
        self.assertEqual(response.status_code, 400)

    def test_ad_login_returns_ai_redirect(self):
        fake_connection = MagicMock()
        with (
            patch.object(portal, 'check_ldap_auth', return_value=(True, None)),
            patch.object(portal, 'get_db_connection', return_value=fake_connection),
            patch.object(portal, 'add_user_to_default_group'),
        ):
            response = self.client.post('/login', json={
                'username': 'ivanov',
                'password': 'secret',
                'return_url': self._return_url('B' * 32),
            })
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])
        self.assertTrue(data['redirect_url'].startswith('https://ai.energoprom.by/sso-login?'))
        self.assertEqual(parse_qs(urlsplit(data['redirect_url']).query)['state'][0], 'B' * 32)

    def test_ai_tile_starts_flow_at_public_ai_address(self):
        with self.client.session_transaction() as session:
            session['logged_in'] = True
            session['username'] = 'ivanov'
        with patch.object(portal, 'can_access_portal_path', return_value=True):
            response = self.client.get('/ai-assistant')
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers['Location'], 'https://ai.energoprom.by')

    def test_healthcheck_is_available_without_login(self):
        fake_connection = MagicMock()
        fake_connection.execute.return_value.fetchone.return_value = (1,)
        with patch.object(portal, 'get_db_connection', return_value=fake_connection):
            response = self.client.get('/healthz')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {'status': 'ok'})
        fake_connection.close.assert_called_once_with()


if __name__ == '__main__':
    unittest.main()
