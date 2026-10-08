import datetime
import unittest
from unittest.mock import patch
import app

class ContractYearTests(unittest.TestCase):
    def setUp(self):
        self.clock = patch.object(app, 'current_date', return_value=datetime.date(2026,10,8))
        self.clock.start()
        self.addCleanup(self.clock.stop)

    def test_current_year_is_year_to_date(self):
        self.assertEqual(app.contract_period(2026), (2026,'2026-01-01','2026-10-08'))
        self.assertEqual(app.contract_period(), app.contract_period(2026))

    def test_previous_year_is_full_year(self):
        self.assertEqual(app.contract_period(2025), (2025,'2025-01-01','2025-12-31'))

    def test_year_forwarded_to_each_company(self):
        with patch.object(app,'mapo_contract_history',return_value={'ok':True}) as history:
            response=app.app.test_client().get('/api/mapo-contracts?year=2025&company=A&company=B')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.json['period'],'2025-01-01 ~ 2025-12-31')
        self.assertEqual({call.args for call in history.call_args_list},{('A',2025),('B',2025)})

    def test_missing_year_defaults_to_current(self):
        with patch.object(app,'mapo_contract_history',return_value={'ok':True}) as history:
            response=app.app.test_client().get('/api/mapo-contracts?company=A')
        history.assert_called_once_with('A',2026)
        self.assertEqual(response.json['period'],'2026-01-01 ~ 2026-10-08')

    def test_invalid_year_rejected_without_network(self):
        with patch.object(app,'mapo_contract_history') as history:
            for year in ('abc','2027','2020','', '2025.5'):
                self.assertEqual(app.app.test_client().get('/api/mapo-contracts?year='+year+'&company=A').status_code,400)
            history.assert_not_called()

    def test_date_conditions_sent_to_official_site(self):
        from types import SimpleNamespace
        with patch.object(app,'get',return_value=SimpleNamespace(text='점검 중',url='https://contract.seoul.go.kr')) as get:
            app.mapo_contract_history('A',2025)
        params=get.call_args.args[1]
        self.assertEqual(params['ps0_fisYear'],'2025')
        self.assertEqual(params['ps0_conYmdS'],'2025-01-01')
        self.assertEqual(params['ps0_conYmdE'],'2025-12-31')
