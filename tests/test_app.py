import unittest
from types import SimpleNamespace
from unittest.mock import patch
import app

HEAD = '<table><tr><th>발주기관</th><th>계약명</th><th>계약업체명</th><th>계약일자</th><th>담당부서</th></tr>'
def page(rows='', total=0):
    return f'<p>총 {total} 건</p>{HEAD}{rows}</table>'
def row(name='테스트', department='총무과', institution='마포구'):
    return f'<tr><td>{institution}</td><td>인쇄</td><td>{name}</td><td>2026-01-01</td><td>{department}</td></tr>'

class ReviewRegressionTests(unittest.TestCase):
    def history(self, html):
        with patch.object(app, 'get', return_value=SimpleNamespace(text=html, url='https://contract.seoul.go.kr/result')):
            return app.mapo_contract_history('테스트', 2026)

    def test_maintenance_is_failure(self):
        self.assertFalse(self.history('<html>서비스 점검 중</html>')['ok'])

    def test_browser_generated_contracts_are_failure(self):
        result = self.history(page('<script class="crypto-data">encrypted</script>'))
        self.assertFalse(result["ok"])
        self.assertIn("브라우저", result["error"])

    def test_count_without_result_table_is_failure(self):
        self.assertFalse(self.history('<p>총 0 건</p>')['ok'])

    def test_explicit_empty_result(self):
        result = self.history(page('<tr><td colspan="5">검색 결과가 없습니다</td></tr>'))
        self.assertTrue(result['ok'])
        self.assertEqual(result['total'], 0)

    def test_exact_company_only(self):
        result = self.history(page(row()+row('테스트상사'), 2))
        self.assertTrue(result['ok'])
        self.assertEqual(result['total'], 1)
        self.assertFalse(result['district_limit'])

    def test_unparsed_count_is_failure(self):
        self.assertFalse(self.history(page(row(), 5))['ok'])

    def test_unrelated_institution_is_failure(self):
        self.assertFalse(self.history(page(row(institution='강남구'), 1))['ok'])

    def test_unknown_department_not_safe_limit(self):
        result = self.history(page(row(department=''), 1))
        self.assertEqual(result['departments'], {'부서 미확인': 1})
        self.assertEqual(result['department_limit'], [])

    def test_regions(self):
        for address in ['마포구', '서울특별시 마포구 월드컵로 1']:
            self.assertTrue(app.matches_region(address, '서울'))
            self.assertTrue(app.matches_region(address, '마포구'))
        for address in ['', '경기도 김포시', '부산광역시 중구']:
            self.assertFalse(app.matches_region(address, '서울'))
        self.assertTrue(app.matches_region('경기도 김포시', '경기'))
        self.assertTrue(app.matches_region('', '전국'))

    def test_board_does_not_fabricate_address(self):
        html='<table><tr><th>업체명</th><th>계약종류</th></tr><tr><td>1</td><td>물품</td><td>인쇄</td><td>제목</td><td>테스트</td><td>123</td><td>2026</td></tr></table>'
        response=SimpleNamespace(text=html,url='https://www.mapo.go.kr',raise_for_status=lambda:None)
        with patch.object(app.requests.Session,'get',return_value=response):
            self.assertEqual(app.mapo_local({})[0]['address'], '')

    def test_residency_and_empty_source_selection(self):
        rows=[dict(source='test',name=str(i),item='',url='',address=address) for i,address in enumerate(['','마포구','서울특별시','서울 강남구'])]
        with patch.object(app,'excel_sources',return_value=(rows,{})), patch.dict(app.ADAPTERS, {'test': lambda q: self.fail('Online source called')}):
            results=app.app.test_client().get('/api/search?sources=').json['results']
        self.assertEqual([x['mapo_local_status'] for x in results], ['확인 필요','관내','확인 필요','관외'])

if __name__ == '__main__':
    unittest.main()
