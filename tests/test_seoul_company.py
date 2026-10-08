import unittest
from types import SimpleNamespace
from unittest.mock import patch
import app


def company_html():
    rows = ''.join(f'''<tr><td colspan="3"><div class="company_sticker">{kind}</div>
        <a><b>주식회사 녹색사람들</b>( 사업자번호 : 119-**-***** )</a></td></tr>
        <tr><td>대표품목 | 기타 제품 제조업</td><td>전화번호 | 02-000-0000</td>
        <td>주소 | 서울 관악구 청룡동</td></tr>''' for kind in ('여성기업', '중기업'))
    return '<p>총 2건</p><table><thead><tr><th>기업명</th><th>주소</th></tr></thead><tbody>'+rows+'</tbody></table>'


class SeoulCompanyTests(unittest.TestCase):
    def test_reads_second_row_and_preserves_types_in_search(self):
        response = SimpleNamespace(text=company_html(), url='https://contract.seoul.go.kr/result')
        with patch.object(app, 'get', return_value=response), patch.object(app, 'excel_sources', return_value=([], {})):
            result = app.app.test_client().get('/api/search', query_string={
                'company': '녹색사람들', 'region': '서울', 'sources': '서울계약마당'}).json
        self.assertEqual(result['count'], 2)
        self.assertEqual({x['type'] for x in result['results']}, {'여성기업', '중기업'})
        for row in result['results']:
            self.assertEqual(row['name'], '주식회사 녹색사람들')
            self.assertEqual(row['address'], '서울 관악구 청룡동')
            self.assertEqual(row['phone'], '02-000-0000')
            self.assertEqual(row['item'], '기타 제품 제조업')
            self.assertEqual(row['mapo_local_status'], '관외')

    def test_missing_detail_is_failure(self):
        response=SimpleNamespace(text=company_html().replace('주소 | 서울 관악구 청룡동', '주소 정보 없음'),url='https://contract.seoul.go.kr/result')
        with patch.object(app,'get',return_value=response):
            with self.assertRaisesRegex(RuntimeError,'상세 행'):
                app.seoul({'company':'녹색사람들'})

    def test_maintenance_is_failure(self):
        with patch.object(app,'get',return_value=SimpleNamespace(text='점검 중',url='https://contract.seoul.go.kr/result')):
            with self.assertRaises(RuntimeError):
                app.seoul({'company':'녹색사람들'})

    def test_valid_zero(self):
        html='<p>총 0건</p><table><thead><tr><th>기업명</th><th>주소</th></tr></thead><tbody></tbody></table>'
        with patch.object(app,'get',return_value=SimpleNamespace(text=html,url='https://contract.seoul.go.kr/result')):
            self.assertEqual(app.seoul({'company':'없는업체'}),[])
