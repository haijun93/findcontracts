from flask import Flask,request,jsonify,send_from_directory
import requests,re,concurrent.futures,urllib.parse,glob,os,datetime
from openpyxl import load_workbook
from bs4 import BeautifulSoup

app=Flask(__name__,static_folder=".")
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/154 Safari/537.36"}

def txt(x): return re.sub(r"\s+"," ",x or "").strip()
def get(url,params=None):
    r=requests.get(url,params=params,headers=H,timeout=18); r.raise_for_status(); return r

def seoul(q):
    u="https://contract.seoul.go.kr/new1/ppviews/searchCompany.do"
    rg=q.get("region","전국")
    p={"searchFlag":"2","recordCountPerPage":"50","currentPageNo":"1","CUS_ADDR1":"0" if rg in ("","전국") else rg,
       "a_CUS_ADDR1":rg or "전국","juso":q.get("address",""),"CUS_NM":q.get("company",""),"CUS_ITM":q.get("keyword","")}
    r=get(u,p)
    soup=BeautifulSoup(r.text,"html.parser")
    table = next((t for t in soup.select("table")
                  if "기업명" in t.get_text() and "주소" in t.get_text()), None)
    if table is None:
        raise RuntimeError("희망기업 결과 테이블 구조를 확인할 수 없습니다")
    out=[]
    for tr in table.select("tbody tr"):
        t=txt(tr.get_text(" ",strip=True))
        m=re.search(r"(.+?)\(\s*사업자번호\s*:\s*([^)]+)\)",t)
        if not m:
            continue
        head=txt(m.group(1))
        sticker=tr.select_one(".company_sticker")
        typ=txt(sticker.get_text(" ",strip=True)) if sticker else ""
        if typ and head.startswith(typ):
            head=head[len(typ):].strip()
        details = tr.find_next_sibling("tr")
        fields = {}
        if details is not None and "사업자번호" not in details.get_text():
            for cell in details.find_all("td", recursive=False):
                label, separator, value = txt(cell.get_text(" ", strip=True)).partition("|")
                if separator:
                    fields[label.strip()] = value.strip()
        if not all(label in fields for label in ("대표품목", "전화번호", "주소")):
            raise RuntimeError("희망기업 상세 행을 해석할 수 없습니다")
        out.append({"source":"서울계약마당","type":typ,"name":head,
                    "item":fields["대표품목"],"phone":fields["전화번호"],
                    "address":fields["주소"],"bizno":txt(m.group(2)),"url":r.url})
    if not out and not re.search(r"총\s*0\s*건", soup.get_text(" ", strip=True)):
        raise RuntimeError("희망기업 검색결과를 해석할 수 없습니다")
    return out


def able(q):
    kw=q.get("keyword") or q.get("company") or ""
    r=get("https://www.ablemarket.or.kr/shop/search.php",{"q":kw,"qbrand":"1","qname":"1","qexplan":"1"})
    soup=BeautifulSoup(r.text,"html.parser"); out=[]
    # Product cards vary; use visible product anchors and nearby brand text.
    for a in soup.select("a"):
        title=txt(a.get_text(" ",strip=True))
        href=a.get("href","")
        if not title or "item.php" not in href: continue
        box=a
        for _ in range(4):
            if box.parent: box=box.parent
        bt=txt(box.get_text(" ",strip=True))
        # best-effort seller/brand: text preceding title
        before=bt.split(title)[0].strip()
        seller=before.split()[-1] if before else ""
        if seller and len(seller)<2: seller=""
        out.append({"source":"에이블마켓","type":"장애인생산품","name":seller or "판매시설 확인",
                    "item":title,"phone":"","address":"","bizno":"",
                    "url":urllib.parse.urljoin(r.url,href)})
    # de-dupe products
    seen=set(); z=[]
    for x in out:
        k=(x["name"],x["item"])
        if k not in seen: seen.add(k);z.append(x)
    return z[:50]

def goods(q):
    # Dreamdre search/list pages are session-sensitive. Query the public product list endpoint.
    kw=q.get("keyword") or q.get("company") or ""
    candidates=[
      ("https://www.goods.go.kr/sm/pd/product/list.do",{"searchKey":"goodsNm","searchValue":kw,"maxPageItems":"50","menuNo":"1200000"}),
      ("https://www.goods.go.kr/sm/pd/product/list.do",{"searchValue":kw,"maxPageItems":"50"})
    ]
    last=None
    for u,p in candidates:
        try:
            r=get(u,p); soup=BeautifulSoup(r.text,"html.parser"); out=[]
            for a in soup.select('a[href*="product/view.do"]'):
                title=txt(a.get_text(" ",strip=True))
                if not title: continue
                href=urllib.parse.urljoin(r.url,a.get("href"))
                parent=a.parent
                block=txt(parent.get_text(" ",strip=True)) if parent else title
                out.append({"source":"꿈드래","type":"중증장애인생산품","name":"생산시설 상세확인",
                            "item":title,"phone":"","address":"","bizno":"","url":href})
            if out:
                seen=set();z=[]
                for x in out:
                    if x["url"] not in seen:seen.add(x["url"]);z.append(x)
                return z[:50]
            last="검색결과 파싱 불가"
        except Exception as e:last=str(e)
    raise RuntimeError(last or "꿈드래 검색 실패")

def generic_form_search(base, q, source, typ):
    """Best-effort official-site form discovery; never uses a third-party search engine."""
    r=get(base); soup=BeautifulSoup(r.text,"html.parser")
    kw=q.get("keyword") or q.get("company") or ""
    forms=soup.find_all("form")
    for f in forms:
        inputs=f.find_all(["input","select"])
        names=[i.get("name") for i in inputs if i.get("name")]
        likely=[n for n in names if any(k in n.lower() for k in ["search","query","keyword","word","text","sch","key"])]
        if not likely: continue
        data={}
        for i in inputs:
            n=i.get("name")
            if n and i.name=="input" and i.get("type","text") in ("hidden","text","search"):
                data[n]=i.get("value","")
        data[likely[-1]]=kw
        action=urllib.parse.urljoin(r.url,f.get("action") or r.url)
        rr=requests.get(action,params=data,headers=H,timeout=18) if (f.get("method","get").lower()=="get") else requests.post(action,data=data,headers=H,timeout=18)
        rr.raise_for_status(); ss=BeautifulSoup(rr.text,"html.parser")
        out=[]
        # collect links whose visible text contains keyword, or surrounding text does
        for a in ss.select("a[href]"):
            label=txt(a.get_text(" ",strip=True))
            if len(label)<2: continue
            block=txt(a.parent.get_text(" ",strip=True)) if a.parent else label
            if kw and kw.lower() not in block.lower(): continue
            if len(block)>500: continue
            out.append({"source":source,"type":typ,"name":label,"item":block,"phone":"","address":"","bizno":"",
                        "url":urllib.parse.urljoin(rr.url,a.get("href"))})
        if out:return out[:50]
    raise RuntimeError("공식 사이트에서 자동 판별 가능한 검색 폼을 찾지 못했습니다.")

def wemall(q): return generic_form_search("https://www.wemall.kr",q,"우리몰","장애인기업")
def social(q): return generic_form_search("https://www.socialenterprise.or.kr",q,"사회적기업진흥원","사회적기업")
def g2b(q): return generic_form_search("https://shopping.g2b.go.kr",q,"나라장터 종합쇼핑몰","조달등록")


def digits(v):
    if v is None: return ""
    if isinstance(v,float) and v.is_integer(): v=int(v)
    return re.sub(r"\D","",str(v))

def newest_matching_file(phrase):
    here=os.path.dirname(os.path.abspath(__file__))
    files=[p for p in glob.glob(os.path.join(here,"*.xlsx")) if phrase in os.path.basename(p) and not os.path.basename(p).startswith("~$")]
    return max(files,key=os.path.getmtime) if files else None

def find_header(ws, required, max_rows=20):
    for i,row in enumerate(ws.iter_rows(min_row=1,max_row=min(max_rows,ws.max_row),values_only=True),1):
        vals=[txt(str(v)) if v is not None else "" for v in row]
        if all(any(req in v for v in vals) for req in required):
            return i, vals
    return None,None


def mapo_excel_rows():
    p=newest_matching_file("마포구 관내 희망기업 목록")
    if not p:
        return [],{"ok":False,"count":0,"error":"같은 폴더에서 파일을 찾지 못함","file":""}
    out=[]
    try:
        wb=load_workbook(p,read_only=True,data_only=True)
        # 1) 중증장애인생산품 생산시설
        for ws in wb.worksheets:
            sn=ws.title
            if "생산시설 지정현황" in sn:
                for row in ws.iter_rows(min_row=4,values_only=True):
                    if len(row)<10 or not row[2]: continue
                    out.append({"source":"Excel·마포희망기업","type":"중증장애인생산품",
                      "name":txt(str(row[2])),"item":txt(str(row[9])) if row[9] is not None else "",
                      "phone":txt(str(row[8])) if row[8] is not None else "",
                      "address":txt(str(row[7])) if row[7] is not None else "",
                      "bizno":digits(row[5]),"url":""})
            elif "사회적기업(" in sn and "예비" not in sn:
                for row in ws.iter_rows(min_row=3,values_only=True):
                    if len(row)<11 or not row[4]: continue
                    item=row[5] if row[5] is not None else (row[7] if len(row)>7 else "")
                    out.append({"source":"Excel·마포희망기업","type":"사회적기업",
                      "name":txt(str(row[4])),"item":txt(str(item)) if item is not None else "",
                      "phone":"","address":txt(str(row[10])) if row[10] is not None else "",
                      "bizno":digits(row[8]),"url":""})
            elif "예비사회적기업" in sn:
                for row in ws.iter_rows(min_row=3,values_only=True):
                    if len(row)<8 or not row[4]: continue
                    out.append({"source":"Excel·마포희망기업","type":"예비사회적기업",
                      "name":txt(str(row[4])),"item":txt(str(row[7])) if row[7] is not None else "",
                      "phone":"","address":"","bizno":digits(row[5]),"url":""})
            elif "사회적협동조합" in sn:
                for row in ws.iter_rows(min_row=5,values_only=True):
                    if len(row)<6 or not row[0]: continue
                    out.append({"source":"Excel·마포희망기업","type":"사회적협동조합",
                      "name":txt(str(row[0])),"item":txt(str(row[5])) if row[5] is not None else "",
                      "phone":txt(str(row[3])) if row[3] is not None else "",
                      "address":txt(str(row[2])) if row[2] is not None else "",
                      "bizno":"","url":""})
        return out,{"ok":True,"count":len(out),"error":"","file":os.path.basename(p)}
    except Exception as e:
        return [],{"ok":False,"count":0,"error":str(e),"file":os.path.basename(p)}

def excel_sources(q):
    files={
      "조달업체 등록 내역":newest_matching_file("조달업체 등록 내역"),
      "중소기업 확인업체 내역":newest_matching_file("중소기업 확인업체 내역")
    }
    bybiz={}; statuses={}
    p=files["조달업체 등록 내역"]
    if p:
        try:
            wb=load_workbook(p,read_only=True,data_only=True); count=0
            for ws in wb.worksheets:
                hr,heads=find_header(ws,["업체명","사업자등록번호"])
                if not hr: continue
                idx={h:i for i,h in enumerate(heads) if h}
                for row in ws.iter_rows(min_row=hr+1,values_only=True):
                    name=txt(str(row[idx["업체명"]])) if idx.get("업체명") is not None and row[idx["업체명"]] is not None else ""
                    if not name: continue
                    biz=digits(row[idx["사업자등록번호"]]) if "사업자등록번호" in idx else ""
                    x=bybiz.setdefault(biz or "N:"+name,{"source":[],"type":[],"name":name,"item":[],"phone":"","address":"","bizno":biz,"url":""})
                    x["source"].append("Excel·조달업체")
                    cg=txt(str(row[idx["기업구분"]])) if "기업구분" in idx and row[idx["기업구분"]] is not None else ""
                    if cg:x["type"].append(cg)
                    x["type"].append("조달등록")
                    x["address"]=txt(str(row[idx["업체소재시군구"]])) if "업체소재시군구" in idx and row[idx["업체소재시군구"]] is not None else x["address"]
                    for h in ("대표업종","대표세부품명"):
                        if h in idx and row[idx[h]] is not None:
                            v=txt(str(row[idx[h]]))
                            if v:x["item"].append(v)
                    count+=1
            statuses["Excel·조달업체"]={"ok":True,"count":count,"error":"","file":os.path.basename(p)}
        except Exception as e: statuses["Excel·조달업체"]={"ok":False,"count":0,"error":str(e),"file":os.path.basename(p)}
    else: statuses["Excel·조달업체"]={"ok":False,"count":0,"error":"같은 폴더에서 파일을 찾지 못함","file":""}

    p=files["중소기업 확인업체 내역"]
    if p:
        try:
            wb=load_workbook(p,read_only=True,data_only=True); count=0
            for ws in wb.worksheets:
                hr,heads=find_header(ws,["업체명","사업자등록번호","기업규모"])
                if not hr: continue
                idx={h:i for i,h in enumerate(heads) if h}
                for row in ws.iter_rows(min_row=hr+1,values_only=True):
                    name=txt(str(row[idx["업체명"]])) if row[idx["업체명"]] is not None else ""
                    if not name: continue
                    biz=digits(row[idx["사업자등록번호"]])
                    x=bybiz.setdefault(biz or "N:"+name,{"source":[],"type":[],"name":name,"item":[],"phone":"","address":"","bizno":biz,"url":""})
                    x["source"].append("Excel·중소기업확인")
                    size=txt(str(row[idx["기업규모"]])) if row[idx["기업규모"]] is not None else ""
                    if size:x["type"].append(size)
                    if "전화번호" in idx and row[idx["전화번호"]] is not None:x["phone"]=txt(str(row[idx["전화번호"]]))
                    if "경쟁입찰참여제한" in idx and row[idx["경쟁입찰참여제한"]] is not None:
                        lim=txt(str(row[idx["경쟁입찰참여제한"]]))
                        if lim and lim!="미해당":x["item"].append("경쟁입찰참여제한: "+lim)
                    count+=1
            statuses["Excel·중소기업확인"]={"ok":True,"count":count,"error":"","file":os.path.basename(p)}
        except Exception as e: statuses["Excel·중소기업확인"]={"ok":False,"count":0,"error":str(e),"file":os.path.basename(p)}
    else: statuses["Excel·중소기업확인"]={"ok":False,"count":0,"error":"같은 폴더에서 파일을 찾지 못함","file":""}

    # 마포구 관내 희망기업 목록도 자동 탐색하여 기존 Excel 자료와 사업자번호 기준으로 결합
    mrows,mstatus=mapo_excel_rows()
    statuses["Excel·마포희망기업"]=mstatus
    for m in mrows:
        key=m.get("bizno") or "N:"+m.get("name","")
        if key in bybiz:
            x=bybiz[key]
            if isinstance(x["source"],list): x["source"].append(m["source"])
            else: x["source"]=[x["source"],m["source"]]
            if isinstance(x["type"],list): x["type"].append(m["type"])
            else: x["type"]=[x["type"],m["type"]]
            if isinstance(x["item"],list):
                if m["item"]: x["item"].append(m["item"])
            else: x["item"]=[x["item"],m["item"]]
            if not x.get("phone"): x["phone"]=m.get("phone","")
            if not x.get("address"): x["address"]=m.get("address","")
        else:
            bybiz[key]={"source":[m["source"]],"type":[m["type"]],"name":m["name"],
                        "item":[m["item"]] if m["item"] else [],"phone":m["phone"],
                        "address":m["address"],"bizno":m["bizno"],"url":""}

    kw=q.get("keyword","").lower(); rg=q.get("region",""); et=q.get("type","").lower(); co=q.get("company","").lower()
    out=[]
    for x in bybiz.values():
        x["source"]=" + ".join(dict.fromkeys(x["source"]))
        x["type"]=" / ".join(dict.fromkeys(x["type"]))
        x["item"]=" / ".join(dict.fromkeys(x["item"]))
        hay=" ".join(str(v) for v in x.values()).lower()
        if kw and kw not in hay: continue
        if co and co not in x["name"].lower(): continue
        if et and et not in hay: continue
        if not matches_region(x["address"], rg): continue
        out.append(x)
    return out,statuses


SEOUL_DISTRICTS = set("종로구 중구 용산구 성동구 광진구 동대문구 중랑구 성북구 강북구 도봉구 노원구 은평구 서대문구 마포구 양천구 강서구 구로구 금천구 영등포구 동작구 관악구 서초구 강남구 송파구 강동구".split())


def address_region(address):
    tokens = re.findall(r"[가-힣]+", address or "")
    aliases = {"서울": "서울", "서울특별시": "서울", "경기": "경기", "경기도": "경기",
               "인천": "인천", "인천광역시": "인천"}
    for token in tokens:
        if token in aliases:
            return aliases[token], tokens
        if token in ("부산", "부산광역시", "대구", "대구광역시", "대전", "대전광역시", "광주", "광주광역시", "울산", "울산광역시", "세종", "세종특별자치시", "충북", "충청북도", "충남", "충청남도", "전북", "전라북도", "전북특별자치도", "전남", "전라남도", "경북", "경상북도", "경남", "경상남도", "강원", "강원도", "강원특별자치도", "제주", "제주특별자치도"):
            return token, tokens
    if tokens and tokens[0] in SEOUL_DISTRICTS - {"중구", "강서구"}:
        return "서울", tokens
    return "", tokens


def matches_region(address, region):
    if not region or region == "전국":
        return True
    province, tokens = address_region(address)
    return province == region if region in ("서울", "경기", "인천") else region in tokens


def company_key(name):
    return re.sub(r"\s+", "", txt(name)).casefold()


def mapo_contract_history(company, year=None):
    """서울계약마당 계약정보에서 해당 업체의 마포구 연간 계약 실적을 조회한다."""
    company=txt(company)
    if not company:
        return {"ok":False,"total":0,"departments":{},"contracts":[],"error":"업체명 없음"}
    today=datetime.date.today()
    year=year or today.year
    start=f"{year}-01-01"
    end=today.isoformat() if year==today.year else f"{year}-12-31"
    base="https://contract.seoul.go.kr/new1/views/contractInfo.do"
    common={
      "ps_selectForm":"0","ps_recordCountPerPage":"50","ps0_fisYear":"",
      "ps0_conTitle":"","ps0_t2OfficeCd":"0","ps0_setOfficeCd":"마포구",
      "ps0_setCustNm":company,"ps0_conYmdS":start,"ps0_conYmdE":end,
      "ps0_conKind":"","ps0_conDiv":"","ps0_conMtdNm":"","ps0_conType":"",
      "ps0_conAmtS":"","ps0_conAmtE":"","ps0_conTitleLogicOperator":"O",
      "ps0_keywordGroup":"","ps_currentPageNo":"1"
    }
    try:
        r=get(base,common)
        soup=BeautifulSoup(r.text,"html.parser")
        page_text=txt(soup.get_text(" ",strip=True))
        mt=re.search(r"총\s*([\d,]+)\s*건",page_text)
        if not mt:
            raise RuntimeError("계약 검색 결과 건수를 확인할 수 없습니다")
        total_hint = int(mt.group(1).replace(",", ""))
        pages = max(1, (total_hint + 49) // 50)
        if pages > 100:
            raise RuntimeError("조회 범위를 초과했습니다. 공식 사이트에서 확인하세요")
        contracts = []
        seen = set()
        parsed_rows = 0
        for pg in range(1, pages + 1):
            p = dict(common); p["ps_currentPageNo"] = str(pg)
            rr = r if pg == 1 else get(base, p)
            ss = BeautifulSoup(rr.text, "html.parser")
            if ss.select("script.crypto-data"):
                raise RuntimeError("계약 목록이 브라우저에서 생성되는 형식입니다. 공식 사이트에서 확인하세요")
            table = None
            columns = {}
            aliases = {"company": ("계약업체명", "계약업체", "업체명", "계약상대자"),
                       "institution": ("발주기관", "기관명", "발주기관명"),
                       "project": ("계약명", "사업명", "계약건명"),
                       "date": ("계약일자", "계약일"),
                       "department": ("담당부서", "발주부서", "부서명")}
            for candidate in ss.select("table"):
                for header in candidate.select("tr"):
                    heads = [txt(c.get_text(" ", strip=True)) for c in header.find_all("th", recursive=False)]
                    cols = {key: next((i for i, h in enumerate(heads) if h in labels), None)
                            for key, labels in aliases.items()}
                    if all(cols[k] is not None for k in ("company", "institution", "project", "date")):
                        table, columns = candidate, cols
                        break
                if table is not None:
                    break
            if table is None:
                raise RuntimeError("계약 결과 테이블 구조를 확인할 수 없습니다")
            for tr in table.select("tr"):
                cells = [txt(c.get_text(" ", strip=True)) for c in tr.find_all("td", recursive=False)]
                if not cells:
                    continue
                if len(cells) == 1 and re.search(r"(?:검색|조회|계약|자료|데이터).*없", cells[0]):
                    continue
                if len(cells) <= max(i for i in columns.values() if i is not None):
                    raise RuntimeError("계약 결과 행을 해석할 수 없습니다")
                seller = cells[columns["company"]]
                institution = cells[columns["institution"]]
                project = cells[columns["project"]]
                date = cells[columns["date"]].replace(".", "-").replace("/", "-").strip("-")
                try:
                    contract_date = datetime.date.fromisoformat(date)
                except ValueError:
                    raise RuntimeError("계약일자를 확인할 수 없습니다")
                link = tr.select_one("a[href]")
                detail_url = urllib.parse.urljoin(rr.url, link["href"]) if link else ""
                if urllib.parse.urlparse(detail_url).scheme not in ("http", "https"):
                    detail_url = ""
                key = (seller, institution, date, project, detail_url)
                if key in seen:
                    continue
                seen.add(key)
                parsed_rows += 1
                if company_key(seller) != company_key(company):
                    continue
                if not re.match(r"^(서울특별시\s*)?마포구(?:$|[\s>/\-])", institution):
                    raise RuntimeError("발주기관이 마포구인지 확인할 수 없습니다")
                if not (datetime.date.fromisoformat(start) <= contract_date <= datetime.date.fromisoformat(end)):
                    raise RuntimeError("조회 기간 밖의 계약이 포함되어 있습니다")
                department = cells[columns["department"]] if columns["department"] is not None else ""
                if not department:
                    department = re.sub(r"^(서울특별시\s*)?마포구", "", institution).strip(" >/-")
                contracts.append({"date": date, "department": department or "부서 미확인",
                                  "project": project, "url": detail_url})
        if parsed_rows != total_hint:
            raise RuntimeError("사이트 건수와 해석한 계약 건수가 달라 완전한 조회를 확인할 수 없습니다")
        depts = {}
        for c in contracts:
            depts[c["department"]] = depts.get(c["department"], 0) + 1
        total = len(contracts)
        dept_limit=[d for d,n in depts.items() if d!="부서 미확인" and n>=2]
        return {"ok":True,"total":total,"departments":depts,"contracts":contracts,
                "district_limit":total>=5,"department_limit":dept_limit,
                "period":f"{start} ~ {end}","url":r.url,"error":""}
    except Exception as e:
        return {"ok":False,"total":0,"departments":{},"contracts":[],"district_limit":False,
                "department_limit":[],"period":f"{start} ~ {end}","url":"","error":str(e)}


# Mapo-gu public company boards. Public HTML layout verified; query form parameters
# are not assumed. We page by board links when discoverable and filter locally.
def mapo_board(q, kind):
    from urllib.parse import urljoin, urlparse, parse_qs
    base='https://www.mapo.go.kr/site/main/entprsInfo/'+('entprsList' if kind=='local' else 'entprsPmarketList')
    session=requests.Session(); session.headers.update(H)
    seen_urls=set(); pending=[base]; found=[]; pages=0
    while pending and pages<12:
        url=pending.pop(0)
        if url in seen_urls: continue
        seen_urls.add(url)
        response=session.get(url,timeout=15);response.raise_for_status()
        soup=BeautifulSoup(response.text,'html.parser');pages+=1
        table=None
        for t in soup.select('table'):
            headers=[txt(x.get_text(' ',strip=True)) for x in t.select('th')]
            if '업체명' in headers and ('대표업종' in headers or '계약종류' in headers):table=t;break
        if table is None:raise RuntimeError('업체 목록 테이블 구조가 변경되었거나 접근이 제한되었습니다')
        headers=[txt(x.get_text(' ',strip=True)) for x in table.select('thead th')]
        if not headers: headers=[txt(x.get_text(' ',strip=True)) for x in table.select('tr th')][:8]
        for tr in table.select('tr'):
            td=tr.find_all('td',recursive=False)
            if not td:continue
            vals=[txt(x.get_text(' ',strip=True)) for x in td]
            if len(vals)<5:continue
            # Documented column order for the two Mapo-gu public lists.
            if kind=='local':
                # 연번, 계약종류, 업종, 제목, 업체명, 전화번호, 작성일
                if len(vals)<7:continue
                name,item,typ,phone=vals[4],vals[2]+' / '+vals[3],vals[1],vals[5]
            else:
                # 연번, 업체명, 기업구분, 대표업종, 대표세부품명, 전화번호
                if len(vals)<6:continue
                name,item,typ,phone=vals[1],vals[3]+' / '+vals[4],vals[2],vals[5]
            if not name or name=='업체명':continue
            a=tr.find('a',href=True)
            detail=urljoin(response.url,a['href']) if a else response.url
            found.append({'source':'마포구 관내업체' if kind=='local' else '마포구 나라장터등록',
                          'type':typ,'name':name,'item':item.strip(' /'),'phone':phone,
                          'address':'','bizno':'','url':detail})
        # Follow only actual pagination links within this board. Limit crawl to 12 pages.
        for a in soup.select('a[href]'):
            label=txt(a.get_text(' ',strip=True))
            href=a.get('href','')
            if not label.isdigit() or not href or href.startswith('javascript:'):continue
            next_url=urljoin(response.url,href)
            if urlparse(next_url).path.rstrip('/')==urlparse(base).path.rstrip('/') and next_url not in seen_urls and next_url not in pending:
                pending.append(next_url)
    keyword=q.get('keyword','').lower().strip();company=q.get('company','').lower().strip();typ=q.get('type','').lower().strip()
    result=[x for x in found if (not keyword or keyword in (x['name']+' '+x['item']).lower()) and (not company or company in x['name'].lower()) and (not typ or typ in x['type'].lower())]
    return result

def mapo_local(q):return mapo_board(q,'local')
def mapo_procurement(q):return mapo_board(q,'procurement')

ADAPTERS={"마포구 관내업체":mapo_local,"마포구 나라장터등록":mapo_procurement,"서울계약마당":seoul,"에이블마켓":able,"꿈드래":goods,"우리몰":wemall,"사회적기업진흥원":social,"나라장터":g2b}

@app.get("/api/search")
def api():
    q=dict(request.args)
    wanted=request.args.get("sources", "").split(",") if "sources" in request.args else list(ADAPTERS)
    rows=[]; status={}
    try:
        exrows,exstatus=excel_sources(q); rows += exrows; status.update(exstatus)
    except Exception as e:
        status["Excel 자동조회"]={"ok":False,"count":0,"error":str(e),"file":""}
    def one(name):
        try:
            z=ADAPTERS[name](q)
            return name,z,None
        except Exception as e:return name,[],str(e)
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
        futs=[ex.submit(one,n) for n in wanted if n in ADAPTERS]
        for f in concurrent.futures.as_completed(futs):
            n,z,e=f.result(); rows+=z; status[n]={"ok":e is None,"count":len(z),"error":e or ""}
    # final client criteria
    rg=q.get("region",""); et=q.get("type",""); co=q.get("company","").lower()
    filtered=[]
    for x in rows:
        hay=" ".join(str(v) for v in x.values()).lower()
        if co and co not in hay: continue
        if not matches_region(x.get("address", ""), rg): continue
        if et and et not in hay: continue
        filtered.append(x)
    # merge exact company records where possible
    merged=[]; seen=set()
    for x in filtered:
        k=(x["source"],x["name"],x.get("type", ""),x["item"],x["url"])
        if k not in seen:seen.add(k);merged.append(x)
    # 관내 여부는 소재지 근거로만 확정한다. 마포구청 목록 출처만으로 소재지를 단정하지 않는다.
    for x in merged:
        province, tokens = address_region(x.get("address", ""))
        if "마포구" in tokens and province == "서울":
            x["mapo_local_status"] = "관내"
            x["mapo_local_reason"] = "소재지에 마포구 표기"
        elif province and (province != "서울" or any(t in SEOUL_DISTRICTS - {"마포구"} for t in tokens)):
            x["mapo_local_status"] = "관외"
            x["mapo_local_reason"] = "소재지가 마포구 외 지역으로 표기"
        else:
            x["mapo_local_status"] = "확인 필요"
            x["mapo_local_reason"] = "소재지 정보 부족 또는 불명확"
    return jsonify(results=merged,count=len(merged),status=status)


@app.get("/api/mapo-contracts")
def mapo_contracts_api():
    names=[txt(x) for x in request.args.getlist("company") if txt(x)]
    if not names:
        one=txt(request.args.get("company",""))
        if one:names=[one]
    # cap to avoid accidentally hammering Seoul site; UI batches visible results.
    names=list(dict.fromkeys(names))[:30]
    out={}
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as ex:
        fs={ex.submit(mapo_contract_history,n):n for n in names}
        for f in concurrent.futures.as_completed(fs):
            n=fs[f]
            try: out[n]=f.result()
            except Exception as e: out[n]={"ok":False,"total":0,"departments":{},"contracts":[],"error":str(e)}
    return jsonify(results=out)

@app.get("/")
def home(): return send_from_directory(".","index.html")
if __name__=="__main__":
    print("통합검색: http://127.0.0.1:5050")
    app.run("127.0.0.1",5050,threaded=True)
