"""Supplemental forms based on example layouts, without carrying over example facts."""
from decimal import Decimal
import re

FORM_LABELS = {
    'authority': '발주기관', 'supplier': '계약업체 상호', 'representative': '업체 대표자',
    'business_number': '사업자등록번호', 'supplier_address': '업체 주소', 'supplier_phone': '업체 연락처',
    'legal_basis': '수의계약 법적 근거', 'contract_reason': '수의계약 선정사유',
    'count_year': '계약횟수 확인 연도', 'department_count': '동일업체 부서 계약횟수',
    'district_count': '동일업체 구 전체 계약횟수', 'count_confirmation': '계약상대자 확인 상태',
    'limit_reason': '계약한도 초과 여부·사유', 'enterprise_type': '업체 유형',
    'competitive_product': '중소기업자간 경쟁제품 해당 여부', 'price_unit': '산출조사 단위',
    'price_quantity': '산출조사 수량', 'price_basis': '산출가격 결정 근거',
    'form_date': '작성·서약 예정일', 'preparer': '조사자·작성자 성명 및 직급',
    'verifier': '확인자 성명 및 직위', 'department_signer': '사업부서 서약자 성명 및 직위',
    'finance_signer': '계약부서 서약자 성명 및 직위', 'ombuds_phone': '옴부즈만 연락처',
    'ombuds_fax': '옴부즈만 팩스',
}
EXTRA_DOCUMENTS = [('price_survey', '산출기초조사서'), ('sole_source', '수의계약사유서'), ('integrity', '청렴계약 이행서약서')]


def validate_forms(value, invalid, error_class, number):
    result = {}
    for key in FORM_LABELS:
        text = value.get(key, '')
        if not isinstance(text, str) or len(text) > 5000 or invalid(text):
            raise error_class('추가 서식 항목은 5천 자 이하의 문자로 입력하세요.')
        result[key] = text.strip()
    for key in ('department_count', 'district_count'):
        if result[key] and not re.fullmatch(r'\d{1,6}', result[key]):
            raise error_class('계약횟수는 0 이상의 정수 또는 빈 값으로 입력하세요.')
    if result['count_year'] and not re.fullmatch(r'20\d{2}', result['count_year']):
        raise error_class('계약횟수 확인 연도는 4자리 연도로 입력하세요.')
    if result['price_quantity'] and (not re.fullmatch(r'\d{1,9}(?:\.\d{1,6})?', result['price_quantity']) or Decimal(result['price_quantity']) <= 0):
        raise error_class('산출조사 수량은 양수 또는 빈 값으로 입력하세요.')
    for key, choices in [('competitive_product', ('', '미확인', '예', '아니오')),
                         ('count_confirmation', ('', '미확인', '확인완료'))]:
        if result[key] not in choices:
            raise error_class(FORM_LABELS[key] + ' 입력값을 확인하세요.')
    rows = value.get('comparisons', [])
    if not isinstance(rows, list) or len(rows) > 20:
        raise error_class('비교견적은 20개 이하로 입력하세요.')
    result['comparisons'] = []
    for row in rows:
        if not isinstance(row, dict): raise error_class('비교견적 형식을 확인하세요.')
        target = {}
        for key in ('vendor', 'note', 'source'):
            v = row.get(key, '')
            if not isinstance(v, str) or len(v) > 1000 or invalid(v):
                raise error_class('비교견적 문자는 1천 자 이하로 입력하세요.')
            target[key] = v.strip()
        target['amount'] = number(row.get('amount'))
        result['comparisons'].append(target)
    return result


def form_warnings(facts, kind):
    warnings = []
    required = {
        'price_survey': ['supplier','price_unit','price_quantity','price_basis','form_date','preparer','verifier'],
        'sole_source': ['supplier','legal_basis','contract_reason','count_year','department_count','district_count','enterprise_type','form_date','preparer','verifier'],
        'integrity': ['authority','form_date','department_signer','finance_signer','ombuds_phone'],
    }
    for key in required.get(kind, []):
        if not facts.get(key): warnings.append(FORM_LABELS[key] + ': 확인 필요')
    if kind == 'price_survey':
        if not facts.get('comparisons'): warnings.append('비교견적이 없습니다. 가격 적정성과 비교견적 필요 여부를 확인하세요.')
        for row in facts.get('comparisons', []):
            if not row['vendor'] or row['amount'] is None:
                warnings.append('비교견적 업체명 또는 금액이 미입력입니다.')
    if kind == 'sole_source':
        if facts.get('count_confirmation') != '확인완료': warnings.append('계약상대자의 계약횟수 확인이 완료되지 않았습니다.')
        if facts.get('competitive_product') not in ('예', '아니오'): warnings.append('중소기업자간 경쟁제품 해당 여부 확인 필요')
        if facts.get('department_count') and facts.get('district_count') and int(facts['department_count']) > int(facts['district_count']):
            warnings.append('부서 계약횟수가 구 전체 계약횟수보다 큽니다. 확인하세요.')
        warnings.append('수의계약 적법성·횟수 한도 충족 여부는 자동 판정하지 않습니다. 적용 기준과 예외를 담당자가 확인하세요.')
    if kind == 'integrity': warnings.append('서명 전 초안입니다. 서약자 본인이 내용을 검토하고 실제 서명 또는 날인해야 합니다.')
    return warnings


def sections(facts, kind, missing, money):
    f = lambda key: facts.get(key) or missing
    if kind == 'price_survey':
        return [('1. 조사대상', [('담당 및 공사명', f'담당: {f("department")}\n공사명: {f("title")}\n단위: {f("price_unit")} / 수량: {f("price_quantity")}')]),
                ('2. 산출가격 및 조사근거', [('산출가격(부가세 포함 여부 확인)', f'산출가격: {money(facts["budget"]["total"])}\n결정 근거: {f("price_basis")}\n비교견적은 동일 범위·수량·부가세 조건인지 확인한다. 업체가 제출한 견적금액과 발주자가 확정하는 산출가격의 일치 여부를 검토한다.')]),
                ('3. 작성·확인', [('작성일 및 확인란', f'{f("form_date")}\n조사자: {f("preparer")} (서명 또는 날인)\n확인자: {f("verifier")} (서명 또는 날인)\n{f("authority")} (분임)재무관 귀하')])]
    if kind == 'sole_source':
        return [('1. 계약개요', [('계약건명 및 예산', f'계약건명: {f("title")}\n소요예산: {money(facts["budget"]["total"])}')]),
                ('2. 수의계약 근거 및 사유', [('근거', f('legal_basis')), ('사유', f('contract_reason'))]),
                ('3. 계약업체', [('업체정보', f'상호: {f("supplier")}\n대표자: {f("representative")}\n사업자등록번호: {f("business_number")}\n주소: {f("supplier_address")}\n연락처: {f("supplier_phone")}')]),
                ('4. 동일업체 계약확인', [('계약횟수', f'확인 연도: {f("count_year")}\n부서: {f("department_count")}회 / 구 전체: {f("district_count")}회\n계약상대자 확인 상태: {f("count_confirmation")}\n한도 초과 여부·사유: {f("limit_reason")}\n※ 계약횟수 미입력은 0회가 아니다. 횟수 기준·예외 및 해당 기간을 별도 확인한다.')]),
                ('5. 업체 현황', [('공공구매 등 확인사항', f'중소기업자간 경쟁제품 해당 여부: {f("competitive_product")}\n기업유형: {f("enterprise_type")}\n증빙의 유효기간과 공사 적용 여부를 담당자가 확인한다.')]),
                ('6. 작성·확인', [('작성일 및 확인란', f'{f("form_date")}\n작성자: {f("preparer")} (서명 또는 날인)\n확인자: {f("verifier")} (서명 또는 날인)')])]
    if kind == 'integrity':
        return [('1. 발주기관 청렴계약 이행서약(안)', [('대상 계약', f'{f("authority")}의 「{f("title")}」 계약에 관한 발주기관 공무원용 서약 초안이다. 실제 사용하는 기관 서식 및 관련 규정을 확인한 뒤 확정한다.'),
                ('공정하고 투명한 집행', '입찰·계약·계약이행 과정(준공 이후 포함)에서 관련 공무원은 관계법령의 절차에 따라 공정하고 투명하게 업무를 집행하고, 공개 대상 정보를 관련 규정에 따라 공개하며 옴부즈만 활동에 협조한다.'),
                ('내부 비리제보자 보호', '내부 비리제보자에게 제보를 이유로 불이익을 주지 않는다.'),
                ('금품·향응 등 수수 금지', '입찰·계약·계약이행 과정에서 관련 공무원은 이유를 불문하고 금품·향응이나 부당한 이익의 제공을 요구하거나 받지 않으며, 이를 위반할 경우 관련 법령에 따른 책임을 진다.')]),
                ('2. 옴부즈만 연락처', [('연락처(현행 여부 확인)', f'전화: {f("ombuds_phone")}\n팩스: {f("ombuds_fax")}')]),
                ('3. 서약일·서약자', [('실제 서명 전 확인', f'{f("form_date")}\n{f("authority")} 사업부서: {f("department_signer")} (서명 또는 날인)\n{f("authority")} 계약부서: {f("finance_signer")} (서명 또는 날인)\n※ 실제 서명·날인을 생성하거나 서약 완료로 표시하지 않는다.')])]
    raise ValueError('Unknown form kind')


def survey_table(facts, missing, money):
    headers = ['견적업체', '견적금액(원)', '조사근거·비고']
    rows = [[facts.get('supplier') or missing, money(facts['budget']['total']), '입력 견적서의 검토 총액 / 부가세 조건 확인']]
    rows.extend([r['vendor'] or missing, money(r['amount']), ' / '.join(v for v in (r['source'], r['note']) if v) or missing] for r in facts.get('comparisons', []))
    return headers, rows
