import tempfile
import unittest
from unittest.mock import MagicMock, patch

import agent
from test_recovery import PRODUCT_HTML, SELLER_HTML, fake_playwright


class DeliveryTests(unittest.TestCase):
    def test_postal_code_match_is_not_a_numeric_substring(self):
        self.assertTrue(agent.delivery_matches('Los Angeles 90001', '90001'))
        self.assertFalse(agent.delivery_matches('Germany', '90001'))
        self.assertFalse(agent.delivery_matches('New York 10001', '90001'))
        self.assertFalse(agent.delivery_matches('900010', '90001'))

    def test_checkpoint_is_not_retried_or_treated_as_delivery_dialog(self):
        page=MagicMock()
        with patch.object(agent,'navigate',return_value=('<html>Continue</html>','CHECKPOINT_REQUIRED')) as nav:
            self.assertEqual(agent.set_delivery(page,'US','90001'),('CHECKPOINT_REQUIRED',''))
        nav.assert_called_once()
        page.locator.assert_not_called()
        page.wait_for_timeout.assert_not_called()

    def test_unconfirmed_location_stops_before_product_requests(self):
        for status,shown in [('BLOCKED_CHALLENGE',''),('UNVERIFIED','Germany'),('OK','New York 10001')]:
            with self.subTest(status=status),tempfile.TemporaryDirectory() as tmp:
                with patch.object(agent,'sync_playwright',fake_playwright()),patch.object(agent,'set_delivery',return_value=(status,shown)),patch.object(agent,'navigate') as nav:
                    with self.assertRaises(agent.ProfileAccessRequired):
                        agent.run([('B000000001','US')],tmp,delay=0,skip_storefront=True)
                nav.assert_not_called()
                state=agent.load_resume(tmp)
                self.assertEqual(state['run_status'],'ACTION_REQUIRED')
                self.assertEqual(state['products'][0]['status'],'NOT_PROCESSED')

    def test_visible_mode_rechecks_delivery_after_user_verification(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(agent,'sync_playwright',fake_playwright()),patch.object(agent,'set_delivery',side_effect=[('BLOCKED_CHALLENGE',''),('OK','Los Angeles 90001')]) as delivery,patch.object(agent,'manual_navigation') as manual,patch.object(agent,'navigate',side_effect=[(PRODUCT_HTML,'HTTP_200'),(SELLER_HTML,'HTTP_200')]):
                agent.run([('B000000001','US')],tmp,headless=False,delay=0,skip_storefront=True)
            manual.assert_called_once()
            self.assertEqual(delivery.call_count,2)
            state=agent.load_resume(tmp)
            self.assertEqual(state['run_status'],'COMPLETED')
            self.assertEqual(state['products'][0]['delivery_zip'],'90001')

    def test_delivery_drift_on_product_page_stops_before_saving_it(self):
        html=PRODUCT_HTML.replace('Los Angeles 90001','Germany')
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(agent,'sync_playwright',fake_playwright()),patch.object(agent,'set_delivery',return_value=('OK','90001')),patch.object(agent,'navigate',return_value=(html,'HTTP_200')) as nav:
                with self.assertRaises(agent.ProfileAccessRequired):
                    agent.run([('B000000001','US')],tmp,delay=0,skip_storefront=True)
            nav.assert_called_once()
            state=agent.load_resume(tmp)
            self.assertEqual(state['run_status'],'ACTION_REQUIRED')
            self.assertEqual(state['products'][0]['status'],'NOT_PROCESSED')
            self.assertTrue(any('DELIVERY_MISMATCH' in e['status'] for e in state['events']))

    def test_resume_with_different_zip_refreshes_listing_and_seller(self):
        product=agent.parse_product(PRODUCT_HTML,'B000000001','US')
        product['delivery_zip']='10001'
        old={'version':1,'seeds':[('B000000001','US')],'products':[product],
             'groups':[{'marketplace':'US','seller_id':'A123456789','seller':agent.parse_seller(SELLER_HTML,'A123456789','US'),
                        'products':{'B000000001':product},'events':[],'catalog_done':True}]}
        html=PRODUCT_HTML.replace('A123456789','A999999999')
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(agent,'sync_playwright',fake_playwright()),patch.object(agent,'set_delivery',return_value=('OK','90001')),patch.object(agent,'navigate',side_effect=[(html,'HTTP_200'),(SELLER_HTML,'HTTP_200')]) as nav:
                agent.run([('B000000001','US')],tmp,delay=0,skip_storefront=True,resume_data=old)
            self.assertEqual(nav.call_count,2)
            state=agent.load_resume(tmp)
            self.assertEqual(state['products'][0]['delivery_zip'],'90001')
            self.assertEqual([g['seller_id'] for g in state['groups']],['A999999999'])


if __name__=='__main__':unittest.main()
