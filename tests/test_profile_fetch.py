import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import agent
import profile_fetch
from test_recovery import PRODUCT_HTML, SELLER_HTML, fake_playwright


class ProfileFetchTests(unittest.TestCase):
    def page(self, url='https://www.amazon.com/sp?seller=A123456789', html=SELLER_HTML):
        page=MagicMock();page.url=url;page.content.return_value=html
        page.title.return_value='Amazon.com'
        return page

    def fetch(self,page,navigate,tmp,**kwargs):
        return profile_fetch.fetch_profile(page,'A123456789','US','https://www.amazon.com/sp?seller=A123456789',navigate,agent.parse_seller,tmp,**kwargs)

    def test_http_200_checkpoint_is_not_a_missing_seller(self):
        page=self.page('https://www.amazon.com/checkpoint/continue?token=secret', '<html>Continue shopping</html>')
        nav=MagicMock(return_value=(page.content(),'HTTP_200'))
        with tempfile.TemporaryDirectory() as tmp:
            seller,events=self.fetch(page,nav,tmp)
            self.assertEqual(seller['status'],'CHECKPOINT_REQUIRED')
            self.assertEqual(seller['contact_status'],'NOT_COLLECTED')
            self.assertEqual(nav.call_count,1)
            page.wait_for_timeout.assert_not_called()
            self.assertTrue(Path(seller['diagnostic_html']).is_file())
            self.assertNotIn('secret',json.dumps(events))
            self.assertNotIn('secret',seller['profile_source'])

    def test_delayed_content_is_waited_for_before_parsing(self):
        page=self.page();nav=MagicMock(return_value=('<html>Loading</html>','HTTP_200'))
        with tempfile.TemporaryDirectory() as tmp:
            seller,events=self.fetch(page,nav,tmp)
        page.locator.return_value.first.wait_for.assert_called_once()
        self.assertEqual(seller['name'],'Example Seller');self.assertEqual(nav.call_count,1)

    def test_unrecognized_profile_has_only_one_retry(self):
        page=self.page(html='<html>Unknown template</html>');nav=MagicMock(return_value=(page.content(),'HTTP_200'))
        with tempfile.TemporaryDirectory() as tmp:
            seller,events=self.fetch(page,nav,tmp)
            self.assertEqual(seller['status'],'PROFILE_UNVERIFIED')
            self.assertEqual(nav.call_count,2)
            self.assertEqual(len(list(Path(tmp).glob('*.html'))),2)

    def test_access_denial_is_not_retried(self):
        page=self.page(html='<html>Access denied</html>');nav=MagicMock(return_value=(page.content(),'HTTP_403'))
        with tempfile.TemporaryDirectory() as tmp:
            seller,_=self.fetch(page,nav,tmp)
        self.assertEqual(seller['status'],'HTTP_403');self.assertEqual(nav.call_count,1)

    def test_manual_verification_waits_for_user_and_reopens_requested_url(self):
        page=self.page('https://www.amazon.com/checkpoint/continue?token=secret')
        def user_completes(_):page.url='https://www.amazon.com/'
        page.wait_for_timeout.side_effect=user_completes
        nav=MagicMock(return_value=(SELLER_HTML,'HTTP_200'))
        html,status=profile_fetch.manual_navigation(page,'https://www.amazon.com/sp?seller=A123456789',nav,'CHECKPOINT_REQUIRED')
        self.assertEqual(status,'HTTP_200');nav.assert_called_once()
        page.click.assert_not_called()

    def test_headless_run_stops_on_first_checkpoint_and_preserves_progress(self):
        factory=fake_playwright()
        page=factory.return_value.__enter__.return_value.chromium.launch.return_value.new_context.return_value.new_page.return_value
        page.url='https://www.amazon.com/checkpoint/continue?token=secret'
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(agent,'sync_playwright',factory),patch.object(agent,'set_delivery',return_value=('OK','10001')),patch.object(agent,'navigate',side_effect=[(PRODUCT_HTML,'HTTP_200'),('','CHECKPOINT_REQUIRED')]) as nav:
                with self.assertRaises(agent.ProfileAccessRequired):
                    agent.run([('B000000001','US'),('B000000002','US')],tmp,delay=0,skip_storefront=True)
            state=agent.load_resume(tmp)
            self.assertEqual(nav.call_count,2)
            self.assertEqual(state['run_status'],'ACTION_REQUIRED')
            self.assertEqual(state['groups'][0]['seller']['status'],'CHECKPOINT_REQUIRED')
            self.assertEqual(state['products'][1]['status'],'NOT_PROCESSED')


if __name__=='__main__':unittest.main()
