# -*- coding: utf-8 -*-
"""
KBID 투찰 결과 확인 자동화 (구글 시트 연동 버전)
파일명: kbid_result_main.py
"""

import time
import re
import os
import gspread
import traceback
from datetime import datetime, timedelta
from urllib.parse import quote_plus, unquote
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, TimeoutException

class KbidConfig:
    """설정값 및 셀렉터 관리"""
    LOGIN_URL = "https://www.kbid.co.kr/login/common_login.htm"
    # 결과 공고 검색을 위해 조금 다른 파라미터를 사용할 수도 있지만 기본 검색도 결과 섹션을 포함함
    SEARCH_URL_TEMPLATE = "https://www.kbid.co.kr/search/index.htm?mid=lge123&txtFindWordTop={}"
    
    CLIENT_SECRETS_FILE = "client_secrets.json"
    TOKEN_FILE = "token.json"
    SPREADSHEET_NAME = "입찰관리"
    
    SELECTORS = {
        "login_check": ["//*[contains(text(), '로그아웃')]", "//a[contains(@href, 'logout')]"],
        "result_section": "//div[contains(@class, 'search_result_wrap')]//div[contains(., '최근 결과공고')]",
        "result_links": "//div[contains(@class, 'search_result_wrap')]//div[contains(., '최근 결과공고')]/following-sibling::div//table//a",
        "result_tab": "//ul[contains(@class, 'tab_bid_detail')]//li[contains(., '개찰결과')]",
        "ranking_table": ["#idCBidTable", ".tbl_ranking", ".tbl_search_list", "//table[contains(., '순위')]"],
        "pagination": "//div[contains(@class, 'paging')]//a"
    }

class GoogleSheetsManager:
    """구글 시트 데이터 입출력 관리"""
    def __init__(self):
        self.client = gspread.oauth(
            credentials_filename=KbidConfig.CLIENT_SECRETS_FILE,
            authorized_user_filename=KbidConfig.TOKEN_FILE
        )
        self.sheet = self.client.open(KbidConfig.SPREADSHEET_NAME)
        self.ws = self.sheet.worksheet("투찰준비")
        self._ensure_result_headers()

    def _ensure_result_headers(self):
        """결과 확인에 필요한 컬럼들이 있는지 확인하고 순서 동기화"""
        current_headers = [h.strip() for h in self.ws.row_values(1)]
        
        # 기본 필드 (결과와 무관한 앞부분)
        base_fields = [
            "투찰상태", "공고번호", "공고명", "지역제한", "입찰개시일", "투찰마감일시", "개찰일시",
            "기초금액", "예가변동폭", "투찰하한율", "계약방법",
            "예상투찰가1", "예상투찰가2", "예상투찰가3"
        ]
        
        # 사용자 요청에 따른 최종 결과 항목 리스트 (정확한 순서)
        result_fields = [
            "참여 업체수", "사정률", "1등 상호명", "1등 업체 입찰금액", "1등 업체 사정률",
            "AIR 채호원 입찰금액", "AIR 채호원 사정률", "AIR 채호원 순위",
            "에어채호원 입찰금액", "에어채호원 사정률", "에어채호원 순위"
        ]
        
        # 1. 기존 헤더에서 결과 관련 필드 및 기본 필드 제거 (나머지 기타 필드 유지 위함)
        all_required = base_fields + result_fields
        others = [h for h in current_headers if h not in all_required and h.replace(" ", "") not in [r.replace(" ", "") for r in all_required]]
        
        # 2. 새로운 헤더 구성 (기본 + 결과 + 기타)
        new_headers = base_fields + result_fields + others
        
        # 3. 변경 사항이 있는지 확인 (단순 순서 변경 포함)
        if current_headers != new_headers:
            print(f"✨ 시트 헤더 순서 및 항목을 업데이트합니다.")
            self.ws.update(values=[new_headers], range_name="A1")

    def get_result_tasks(self):
        """개찰일시가 지났고 낙찰확인이 안 된 공고 목록 가져오기"""
        all_data = self.ws.get_all_records()
        now = datetime.now()
        tasks = []
        
        for idx, row in enumerate(all_data, start=2):
            status = str(row.get("투찰상태", "")).strip()
            if status == "낙찰확인":
                continue
                
            open_time_str = str(row.get("개찰일시", "")).strip()
            open_time = self._parse_datetime(open_time_str)
            
            if open_time and now > open_time:
                tasks.append({
                    "row_idx": idx,
                    "name": row.get("공고명", ""),
                    "num": row.get("공고번호", "")
                })
        return tasks

    def _parse_datetime(self, text):
        if not text: return None
        try:
            # 2024-05-10 10:00 형태 처리
            return datetime.strptime(text[:16], "%Y-%m-%d %H:%M")
        except:
            try:
                # 24.05.10 10:00 형태 등 다양한 시도 (정규식 활용)
                parts = re.findall(r'\d+', text)
                if len(parts) >= 5:
                    y, m, d, h, mi = map(int, parts[:5])
                    if y < 100: y += 2000
                    return datetime(y, m, d, h, mi)
            except: pass
        return None

    def update_row(self, row_idx, data_dict):
        """특정 행의 결과 데이터 업데이트"""
        headers = self.ws.row_values(1)
        # 전체 행 데이터를 가져와서 필요한 부분만 수정
        current_row = self.ws.row_values(row_idx)
        new_row = list(current_row)
        
        # 만약 current_row가 헤더보다 짧으면 패딩
        if len(new_row) < len(headers):
            new_row.extend([""] * (len(headers) - len(new_row)))

        for key, value in data_dict.items():
            if key in headers:
                idx = headers.index(key)
                new_row[idx] = value
        
        # A{row_idx}:Z{row_idx} 형태의 범위 계산
        last_col = gspread.utils.rowcol_to_a1(row_idx, len(headers))
        range_name = f"A{row_idx}:{last_col}"
        self.ws.update(values=[new_row], range_name=range_name)

class KbidBrowser:
    def __init__(self):
        self.driver = self._init_driver()

    def _init_driver(self):
        options = Options()
        options.add_argument("--incognito")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument('--disable-blink-features=AutomationControlled')
        driver = webdriver.Chrome(options=options)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        driver.implicitly_wait(3)
        return driver

    def login(self):
        print("🔑 로그인을 확인합니다...")
        self.driver.get(KbidConfig.LOGIN_URL)
        
        # 수동 로그인 대기 (최대 3분)
        start_time = time.time()
        while time.time() - start_time < 180:
            try:
                # 로그인 상태 확인 (로그아웃 버튼 존재 여부 등)
                is_logged_in = "login" not in self.driver.current_url.lower() or \
                              self.driver.find_elements(By.XPATH, "//*[contains(text(), '로그아웃')]")
                
                if is_logged_in:
                    print("✅ 로그인 성공")
                    return True
                time.sleep(2)
            except UnexpectedAlertPresentException as e:
                alert_text = str(e.alert_text) if e.alert_text else "알 수 없는 알림"
                print(f"\n⚠️ 알림 발생: {alert_text}")
                print("   [안내] 브라우저에서 알림창의 '확인' 버튼을 클릭해 주세요. 그 후 작업을 계속합니다.")
                # 알림이 사라질 때까지 대기
                while True:
                    try:
                        time.sleep(2)
                        self.driver.title # 알림이 있으면 여기서 예외 발생
                        break
                    except UnexpectedAlertPresentException:
                        continue
                    except: break
            except Exception as e:
                time.sleep(2)
                
        print("❌ 로그인 대기 시간 초과")
        return False

    def navigate_to_result_bid(self, task):
        """검색 후 '최근 결과공고' 영역에서 클릭"""
        search_term = task["num"] if task["num"] else task["name"]
        match = re.search(r'[A-Z0-9]{5,}-[A-Z0-9]+', search_term)
        clean_num = match.group() if match else search_term
        
        # 검색어 정제
        clean_term = re.sub(r'^(?:결|전|수|취|긴|견|재)\s*', '', clean_num).strip()
        url = KbidConfig.SEARCH_URL_TEMPLATE.format(quote_plus(clean_term))
        
        # [최적화] 메인페이지 정체 방지 및 중복 로드 방지
        for attempt in range(2):
            try:
                current_url = unquote(self.driver.current_url)
                # 이미 해당 검색 결과 페이지에 있는지 확인
                if clean_term in current_url and "search/index.htm" in current_url:
                    print("   [디버그] 이미 해당 검색 결과 페이지에 있습니다.")
                    break
                
                print(f"   [디버그] 검색 페이지 이동 시도 ({attempt+1}/2): {clean_term}")
                self.driver.set_page_load_timeout(20)
                self.driver.get(url)
                time.sleep(1.5)
                
                # 검색 결과 페이지의 핵심 요소가 나타났는지 확인
                # (메인페이지에 머물고 있으면 이 조건이 실패함)
                WebDriverWait(self.driver, 7).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'search_result_wrap') or contains(@class, 'area_result') or @id='idCBidTable' or contains(@id, 'listBody')]"))
                )
                break
            except Exception as e:
                if attempt == 0:
                    print(f"   ⚠️ 검색 페이지 진입 지연 또는 실패 ({e}) - 재시도합니다.")
                    try: self.driver.execute_script("window.stop();")
                    except: pass
                    time.sleep(1)
                    continue
                else:
                    print("   ⚠️ 검색 결과 로딩 지연 중 (계속 진행)")
        
        try:
            # 디버깅용: 항상 현재 검색 결과 저장 (사용자 요청)
            with open("search_result_debug.html", "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
                
            # 1. '결과공고' 섹션 타이틀 찾기
            xpath_title = "//*[contains(text(), '결과공고')]"
            titles = self.driver.find_elements(By.XPATH, xpath_title)
            
            target_link = None
            print(f"   [디버그] 결과공고 관련 타이틀 {len(titles)}개 발견")
            
            # 검색어에서 불필요한 공백 제거
            match_term = clean_num.replace(" ", "")
            # 국방부 공고 여부 확인
            is_mnd = "국방부" in str(task.get("num", ""))
            
            for title_elem in titles:
                try:
                    title_text = title_elem.text.strip()
                    if "입찰" in title_text and "결과" not in title_text:
                        continue # 입찰공고 타이틀은 스킵
                        
                    print(f"   [디버그] 타이틀 '{title_text}' 하위 탐색 시작 (국방부 여부: {is_mnd})")
                    
                    # 타이틀 이후의 모든 tr 탐색 (다음 섹션 타이틀 전까지만)
                    # 1. 타이틀 이후의 모든 tr을 가져옴
                    rows = title_elem.find_elements(By.XPATH, "./following::tr[position() <= 20]")
                    
                    for idx, row in enumerate(rows):
                        try:
                            # [수정] 국방부 공고는 번호 체계가 다르므로 결과 섹션의 첫 번째 행을 우선 선택
                            if is_mnd and idx == 0:
                                links = row.find_elements(By.TAG_NAME, "a")
                                if links:
                                    target_link = links[0]
                                    print(f"   ✅ [국방부] 첫 번째 결과 항목 자동 매칭")
                                    break

                            # 행의 전체 텍스트 (공백 제거)
                            row_text = (row.get_attribute("innerText") or row.text).replace(" ", "").replace("\n", "")
                            
                            # 공고번호가 포함되어 있는지 확인
                            if match_term in row_text:
                                # 해당 행 내의 모든 링크 확인
                                links = row.find_elements(By.TAG_NAME, "a")
                                for link in links:
                                    link_text = link.text.strip()
                                    if link_text: # 텍스트가 있는 링크가 보통 공고 상세 링크
                                        target_link = link
                                        print(f"   ✅ '{title_text}' 영역에서 '{match_term}' 매칭 성공")
                                        break
                        except: continue
                        if target_link: break
                    if target_link: break
                except: continue
            
            # 2. 전수 조사 (최후의 수단: 페이지 전체에서 '결' 배지가 있는 행 탐색)
            if not target_link:
                print("   [디버그] 섹션 기반 탐색 실패, 페이지 전체 전수 조사 시작...")
                # 모든 tr을 가져와서 '결' 아이콘과 번호가 동시에 있는 행 찾기
                all_rows = self.driver.find_elements(By.TAG_NAME, "tr")
                for row in all_rows:
                    try:
                        row_html = row.get_attribute("innerHTML")
                        if 'alt="결"' in row_html:
                            row_text = (row.get_attribute("innerText") or row.text).replace(" ", "")
                            if match_term in row_text:
                                target_link = row.find_element(By.TAG_NAME, "a")
                                print(f"   ✅ 페이지 전수 조사(결 배지)로 결과 항목 발견")
                                break
                    except: continue

            if target_link:
                self.driver.execute_script("arguments[0].click();", target_link)
                # 새 창 전환
                WebDriverWait(self.driver, 10).until(lambda d: len(d.window_handles) > 1)
                self.driver.switch_to.window(self.driver.window_handles[-1])
                return True
            else:
                if "검색결과가 없습니다" in self.driver.page_source:
                    print(f"   ⚠️ 검색 결과가 존재하지 않습니다: {clean_term}")
                else:
                    print(f"   ⚠️ '{clean_num}'에 대한 결과 공고 매칭 실패")
                return False
        except Exception as e:
            print(f"⚠️ 검색 진입 중 오류: {e}")
            return False

class KbidParser:
    def __init__(self, driver):
        self.driver = driver

    def _format_amount(self, text):
        """금액 텍스트에서 숫자만 추출하여 1,000 단위 표시 (원 제거)"""
        if not text: return text
        if text == "-": return text
        
        # 1. 숫자만 추출 (콤마, 한글, 특수문자 제거)
        digits = re.sub(r"[^0-9]", "", str(text))
        
        if not digits: return text
        
        try:
            # 2. 정수로 변환 후 콤마 추가
            val = int(digits)
            return format(val, ',')
        except:
            return text

    def _format_rate(self, text):
        """사정률 등에서 % 제거"""
        if not text: return text
        if text == "-": return text
        return str(text).replace("%", "").strip()

    def verify_result_page(self):
        """'개찰결과' 화면인지 확인하고 필요시 탭 클릭"""
        # 디버깅용: 상세 페이지 소스 저장
        try:
            with open("detail_page_debug.html", "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
        except: pass

        # 1. 페이지 내에 '낙찰순위' 또는 '개찰결과'라는 텍스트가 큰 제목으로 있는지 확인 (이미 진입했을 가능성)
        body_text = self.driver.page_source
        if "낙찰순위" in body_text or "참여업체" in body_text:
            print("   [디버그] 이미 개찰결과 데이터가 화면에 보입니다.")
            return True

        # 2. 탭 클릭 시도
        try:
            # 여러 형태의 탭 셀렉터 시도
            tab_xpaths = [
                KbidConfig.SELECTORS["result_tab"],
                "//li[contains(., '개찰결과')]",
                "//a[contains(., '개찰결과')]",
                "//span[contains(., '개찰결과')]"
            ]
            
            for xp in tab_xpaths:
                tabs = self.driver.find_elements(By.XPATH, xp)
                for tab in tabs:
                    if tab.is_displayed():
                        # 이미 활성화된 탭인지 확인
                        cls = tab.get_attribute("class") or ""
                        if "on" in cls or "active" in cls:
                            return True
                        self.driver.execute_script("arguments[0].click();", tab)
                        time.sleep(2)
                        return True
            return False
        except:
            return False

    def parse_full_results(self):
        """결과 데이터 추출 (참여업체, 사정률, 1등, AIR/에어 등)"""
        data = {
            "참여 업체수": "", "사정률": "", "1등 상호명": "",
            "1등 업체 입찰금액": "", "1등 업체 사정률": "",
            "AIR 채호원 입찰금액": "-", "AIR 채호원 사정률": "-", "AIR 채호원 순위": "-",
            "에어채호원 입찰금액": "-", "에어채호원 사정률": "-", "에어채호원 순위": "-"
        }
        
        # 요약 정보 (참여 업체수, 사정률)
        data["사정률"] = self._format_rate(self._find_text_by_label("사정률") or self._find_text_by_label("낙찰율"))
        
        # 1. 참여 업체수 파싱
        try:
            body_text = self.driver.page_source
            # 사용자 제보 형식: 참여업체수 : 937 [32 / 1 ]
            match = re.search(r"참여업체(?:수)?\s*[:：]\s*([\d,]+)", body_text)
            if match:
                data["참여 업체수"] = match.group(1).replace(",", "")
        except: pass

        # 2. 1등 정보 및 기본 테이블 데이터 (첫 페이지)
        headers, rows = self._get_table_data()
        if rows:
            first_row = rows[0]
            data["1등 상호명"] = self._get_cell(first_row, headers, "상호")
            data["1등 업체 입찰금액"] = self._format_amount(self._get_cell(first_row, headers, "입찰금액"))
            data["1등 업체 사정률"] = self._format_rate(self._get_cell(first_row, headers, "사정률"))

        # 3. 채호원 검색 (AIR/에어 정보 추출용)
        self._search_and_parse_target_companies(data)
        
        return data

    def _search_and_parse_target_companies(self, data):
        """URL 조작을 통한 '채호원' 검색 및 데이터 추출"""
        try:
            current_url = self.driver.current_url
            if "txtResultSearchWord" in current_url: return # 이미 검색 중이면 무시
            
            # 검색 파라미터 추가
            search_param = "&lstResultFields=ComName&txtResultSearchWord=" + quote_plus("채호원")
            search_url = current_url + search_param
            
            print(f"   🔍 '채호원' 검색 페이지로 이동 중...")
            self.driver.get(search_url)
            time.sleep(2)
            
            headers, rows = self._get_table_data()
            found_air = False
            found_corp = False
            
            for row in rows:
                name = self._get_cell(row, headers, "상호").replace(" ", "")
                if "AIR채호원" in name and not found_air:
                    data["AIR 채호원 입찰금액"] = self._format_amount(self._get_cell(row, headers, "입찰금액"))
                    data["AIR 채호원 사정률"] = self._format_rate(self._get_cell(row, headers, "사정률"))
                    data["AIR 채호원 순위"] = self._get_cell(row, headers, "순위")
                    found_air = True
                if ("에어채호원" in name or "애어체호원" in name) and not found_corp:
                    data["에어채호원 입찰금액"] = self._format_amount(self._get_cell(row, headers, "입찰금액"))
                    data["에어채호원 사정률"] = self._format_rate(self._get_cell(row, headers, "사정률"))
                    data["에어채호원 순위"] = self._get_cell(row, headers, "순위")
                    found_corp = True
                if found_air and found_corp: break
            
            if found_air or found_corp:
                print(f"   ✅ 대상 업체 발견: AIR({data['AIR 채호원 순위']}위), 에어({data['에어채호원 순위']}위)")
            else:
                print("   ⚠️ 검색 결과 내에 대상 업체(채호원)가 없습니다.")
        except Exception as e:
            print(f"   ⚠️ 채호원 검색 중 오류: {e}")

    def _find_text_by_label(self, label):
        try:
            xpath = f"//th[contains(., '{label}')]/following-sibling::td[1]"
            return self.driver.find_element(By.XPATH, xpath).text.strip()
        except: return ""


    def _get_table_data(self):
        """다양한 셀렉터로 테이블을 시도하고 데이터 반환"""
        table = None
        selectors = KbidConfig.SELECTORS["ranking_table"]
        if isinstance(selectors, str): selectors = [selectors]
        
        for sel in selectors:
            try:
                if sel.startswith("//") or sel.startswith("("):
                    table = self.driver.find_element(By.XPATH, sel)
                elif sel.startswith("#"):
                    table = self.driver.find_element(By.ID, sel[1:])
                else:
                    table = self.driver.find_element(By.CSS_SELECTOR, sel)
                if table.is_displayed(): break
            except: continue
            
        if not table: return [], []

        try:
            header_elements = table.find_elements(By.XPATH, ".//th")
            headers = [h.text.strip() for h in header_elements]
            
            rows = []
            tr_elements = table.find_elements(By.XPATH, ".//tr[td]")
            for tr in tr_elements:
                cells = [c.text.strip() for c in tr.find_elements(By.TAG_NAME, "td")]
                if len(cells) >= len(headers):
                    rows.append(cells)
            return headers, rows
        except: return [], []

    def _get_cell(self, row, headers, keyword):
        for i, h in enumerate(headers):
            if keyword in h:
                return row[i] if i < len(row) else ""
        return ""

    def _go_to_next_page(self, next_page_num):
        try:
            # 페이지네이션 링크 찾기 (숫자 2, 3... 또는 '다음' 버튼)
            # 숫자로 된 링크 우선 시도
            links = self.driver.find_elements(By.XPATH, KbidConfig.SELECTORS["pagination"])
            for link in links:
                if link.text.strip() == str(next_page_num):
                    self.driver.execute_script("arguments[0].click();", link)
                    return True
            # 숫자가 없으면 '다음' 버튼(보통 > 또는 [다음]) 시도
            for link in links:
                if ">" in link.text or "다음" in link.text:
                    self.driver.execute_script("arguments[0].click();", link)
                    return True
        except: pass
        return False

class KbidResultCrawler:
    def __init__(self):
        self.gs = GoogleSheetsManager()
        self.browser = KbidBrowser()
        self.parser = KbidParser(self.browser.driver)

    def run(self):
        try:
            self.browser.login()
            tasks = self.gs.get_result_tasks()
            print(f"📝 총 {len(tasks)}건의 결과 확인 작업을 시작합니다.")
            
            for task in tasks:
                print(f"\n🔍 [{task['num']}] {task['name']} 처리 중...")
                if self.browser.navigate_to_result_bid(task):
                    if self.parser.verify_result_page():
                        result_data = self.parser.parse_full_results()
                        
                        # 가이드라인: 모든 페이지를 확인했으나 대상 업체가 없다면 '확인불가'로 처리
                        # (단, 1등 정보는 수집된 상태일 수 있음)
                        if result_data.get("AIR 채호원 순위") == "-" and result_data.get("에어채호원 순위") == "-":
                            print("⚠️ 대상 업체를 찾을 수 없어 '확인불가'로 설정합니다.")
                            result_data["투찰상태"] = "확인불가"
                        else:
                            result_data["투찰상태"] = "낙찰확인"
                        
                        self.gs.update_row(task["row_idx"], result_data)
                        if result_data["1등 상호명"]:
                            print(f"✅ 데이터 추출 완료: 1등={result_data['1등 상호명']}, AIR={result_data.get('AIR 채호원 순위', 'X')}, 에어={result_data.get('에어채호원 순위', 'X')}")
                    else:
                        print("❌ 개찰결과 탭을 찾을 수 없습니다. (잘못된 진입)")
                        self.gs.update_row(task["row_idx"], {"투찰상태": "확인불가"})
                    
                    # 상세 창 닫기
                    if len(self.browser.driver.window_handles) > 1:
                        self.browser.driver.close()
                        self.browser.driver.switch_to.window(self.browser.driver.window_handles[0])
                else:
                    print("❌ 결과 공고를 찾을 수 없습니다.")
                    self.gs.update_row(task["row_idx"], {"투찰상태": "확인불가"})
                
                time.sleep(1)
        except Exception as e:
            print(f"🛑 치명적 오류: {e}")
            traceback.print_exc()
        finally:
            self.browser.driver.quit()
            print("\n🏁 모든 작업을 마쳤습니다.")

if __name__ == "__main__":
    KbidResultCrawler().run()
