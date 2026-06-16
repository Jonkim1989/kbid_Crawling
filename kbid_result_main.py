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
    MAIN_URL = "https://www.kbid.co.kr/"
    SEARCH_URL = "https://www.kbid.co.kr/search/index.htm"
    # 결과 공고 검색을 위해 조금 다른 파라미터를 사용할 수도 있지만 기본 검색도 결과 섹션을 포함함
    SEARCH_URL_TEMPLATE = "https://www.kbid.co.kr/search/index.htm?mid=lge123&txtFindWordTop={}"
    
    CLIENT_SECRETS_FILE = "client_secrets.json"
    TOKEN_FILE = "token.json"
    SPREADSHEET_NAME = "입찰관리"
    DEBUG_HISTORY_FOLDER = "debug_history"
    
    SELECTORS = {
        "login_check": ["//*[contains(text(), '로그아웃')]", "//a[contains(@href, 'logout')]"],
        "result_section": "//div[contains(@class, 'search_result_wrap')]//div[contains(., '최근 결과공고')]",
        "result_links": "//div[contains(@class, 'search_result_wrap')]//div[contains(., '최근 결과공고')]/following-sibling::div//table//a",
        "result_tab": "//ul[contains(@class, 'tab_bid_detail')]//li[contains(., '개찰결과')]",
        "ranking_table": ["#idCBidTable", ".tbl_ranking", ".tbl_search_list", "//table[.//th[text()='순위']]"],
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

    def _extract_bid_base_and_serial(self, bid_no):
        """공고번호를 기본 부분과 일련번호로 분해
        예: 'R26BK01526419-002' -> ('R26BK01526419', 2)
        예: 'R26BK01526419' -> ('R26BK01526419', -1)
        """
        bid_no = str(bid_no).strip()
        # 기본 부분과 일련번호 분리 (대시로 분리)
        if '-' in bid_no:
            parts = bid_no.rsplit('-', 1)  # 마지막 대시로 분리
            base = parts[0].strip()
            try:
                serial = int(parts[1].strip())
            except ValueError:
                # 일련번호가 숫자가 아니면 전체를 기본으로 취급
                base = bid_no
                serial = -1
        else:
            base = bid_no
            serial = -1  # 일련번호가 없음을 표시
        return base, serial

    def _ensure_result_headers(self):
        """결과 확인에 필요한 컬럼들이 있는지 확인하고 순서 동기화"""
        current_headers = [h.strip() for h in self.ws.row_values(1)]
        
        # kbid_Crawling_main.py와 동일한 순서로 정렬
        desired_headers = [
            "투찰상태", "공고명", "공고번호", "공고기관", "지역제한", "업종", "입찰개시일", "투찰마감일시", "개찰일시",
            "기초금액", "A값", "예가변동폭", "투찰하한율", "계약방법",
            "예상투찰가1", "예상투찰가2", "예상투찰가3",
            "참여 업체수", "사정률", "1등 상호명", "1등 업체 입찰금액", "1등 업체 사정률",
            "AIR 채호원 입찰금액", "AIR 채호원 사정률", "AIR 채호원 순위",
            "에어채호원 입찰금액", "에어채호원 사정률", "에어채호원 순위"
        ]
        
        # 변경 사항이 있는지 확인 (단순 순서 변경 포함)
        if current_headers != desired_headers:
            print(f"✨ 시트 헤더 순서 및 항목을 업데이트합니다.")
            self.ws.update(values=[desired_headers], range_name="A1")

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
        options.page_load_strategy = 'eager'  # DOM 준비되면 즉시 진행 (외부 트래커 무한대기 방지)
        options.add_argument("--incognito")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument('--disable-blink-features=AutomationControlled')
        driver = webdriver.Chrome(options=options)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        driver.set_page_load_timeout(30)  # eager 모드에서는 DOM 준비 후 타임아웃
        driver.implicitly_wait(3)
        return driver

    def login(self):
        """수동 로그인 대기 (10초 카운트다운 후 사용자 확인)
        - 10초 카운트다운
        - 사용자가 로그인 완료했는지 터미널에서 확인
        - y 입력 시 검색 페이지로 이동
        """
        print("\n🔑 로그인을 확인합니다. (10초 안에 로그인해주세요)\n")
        self.driver.get(KbidConfig.LOGIN_URL)
        
        start_time = time.time()
        timeout = 10
        
        while time.time() - start_time < timeout:
            try:
                elapsed = int(time.time() - start_time)
                remaining = timeout - elapsed
                
                # 남은 시간을 백분율로 표시
                progress_bar = "█" * elapsed + "░" * remaining
                print(f"\r⏳ [{progress_bar}] {remaining}초 남음", end="", flush=True)
                
                time.sleep(0.5)
                
            except UnexpectedAlertPresentException as e:
                alert_text = str(e.alert_text) if e.alert_text else "알 수 없는 알림"
                print(f"\n⚠️ 알림 발생: {alert_text}")
                print("   [안내] 브라우저에서 알림창의 '확인' 버튼을 클릭해 주세요.")
                # 알림이 사라질 때까지 대기
                while True:
                    try:
                        time.sleep(0.5)
                        self.driver.title
                        break
                    except UnexpectedAlertPresentException:
                        continue
                    except: 
                        break
            except Exception as e:
                time.sleep(0.5)
        
        # 카운트다운 완료 (0초 도달)
        print("\n")
        
        # 사용자 확인 루프
        while True:
            user_input = input("✅ 로그인을 완료하셨습니까? (y/n): ").strip().lower()
            
            if user_input == 'y':
                print("🔄 검색 페이지로 이동 중...\n")
                # [개선] 로그인 후 공고 검색 페이지로 이동을 2회 반복
                # (메인페이지 진입 지연 문제 해결)
                for attempt in range(2):
                    time.sleep(3)
                    print(f"   [공고 검색 페이지 이동] {attempt+1}/2 시도...")
                    try:
                        self.driver.get(KbidConfig.SEARCH_URL)
                        time.sleep(10)
                    except Exception as e:
                        print(f"   ⚠️ 이동 중 오류: {str(e)[:100]}")
                        try:
                            self.driver.execute_script("window.stop();")
                        except:
                            pass
                
                return True
            elif user_input == 'n':
                print("⏳ 로그인을 완료한 후 다시 입력해주세요.")
                continue
            else:
                print("⚠️ y 또는 n을 입력해주세요.")
                continue

    def navigate_to_result_bid(self, task):
        """검색 후 '최근 결과공고' 영역에서 클릭"""
        search_term = str(task["num"]) if task["num"] else task["name"]
        match = re.search(r'[A-Z0-9]{5,}-[A-Z0-9]+', search_term)
        clean_num = match.group() if match else search_term
        
        # 검색어 정제
        clean_term = re.sub(r'^(?:결|전|수|취|긴|견|재)\s*', '', clean_num).strip()
        url = KbidConfig.SEARCH_URL_TEMPLATE.format(quote_plus(clean_term))
        
        # [최적화] 메인페이지 정체 방지 및 중복 로드 방지 (강력한 탈출 로직 추가)
        for attempt in range(3): # 시도 횟수 3회로 증가
            try:
                # 0. 윈도우 상태 관리 (불필요한 팝업창 닫기 및 메인창 집중)
                if len(self.driver.window_handles) > 1:
                    main_handle = self.driver.window_handles[0]
                    for handle in self.driver.window_handles[1:]:
                        self.driver.switch_to.window(handle)
                        self.driver.close()
                    self.driver.switch_to.window(main_handle)

                # 1. 현재 상태 확인
                try:
                    current_url = unquote(self.driver.current_url)
                except:
                    print("   ⚠️ 브라우저 응답 없음 - 잠시 대기...")
                    time.sleep(2)
                    continue

                # 이미 해당 검색 결과 페이지에 있는지 확인
                if clean_term in current_url and "search/index.htm" in current_url:
                    print("   [디버그] 이미 해당 검색 결과 페이지에 있습니다.")
                    break
                
                print(f"   [디버그] 검색 페이지 이동 시도 ({attempt+1}/3): {clean_term}")
                
                # 2. 페이지 이동 (eager 모드: DOM 미로딩 자원 있어도 진행)
                if "index_first" in current_url or current_url.endswith(".co.kr/"):
                    search_input = self.driver.find_elements(By.ID, "s_search_word")
                    if search_input and search_input[0].is_displayed():
                        print("   [디버그] 메인페이지 검색창 직접 입력 시도")
                        search_input[0].clear()
                        search_input[0].send_keys(clean_term)
                        search_input[0].send_keys("\n")
                    else:
                        self.driver.get(url)
                else:
                    try:
                        self.driver.get(url)
                    except Exception:
                        # eager 모드에서는 실제 발생하지 않지만, 혁시 나오면 데이터는 이미 있음
                        try: self.driver.execute_script("window.stop();")
                        except: pass
                
                # 3. ViewBid 링크가 있는 행이 나타날 때까지 최대 10초 대기
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//tr[.//a[contains(@href,'ViewBid')]]"))
                )
                time.sleep(2.0)  # AJAX 추가 로딩 여유 시간
                print("   ✅ 검색 결과 페이지 진입 성공")
                break
            except Exception as e:
                print(f"   ⚠️ 검색 페이지 진입 지연 ({attempt+1}/3): {str(e)[:100]}")
                try:
                    self.driver.execute_script("window.stop();")
                    # 알림창이 있을 수 있으므로 확인
                    self.driver.switch_to.alert.accept()
                except: pass
                time.sleep(2)
                if attempt == 2:
                    print("   ⚠️ 최종 페이지 로드 지연 (무시하고 진행)")
        
        try:
            # 디버깅용: 항상 현재 검색 결과 저장 (사용자 요청)
            os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
            with open(f"{KbidConfig.DEBUG_HISTORY_FOLDER}/search_result_debug.html", "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
                
            # 1. '결과공고' 섹션 타이틀 찾기
            xpath_title = "//*[contains(text(), '결과공고')]"
            titles = self.driver.find_elements(By.XPATH, xpath_title)
            
            target_link = None
            print(f"   [디버그] 결과공고 관련 타이틀 {len(titles)}개 발견")
            
            # 매칭용 텍스트 정제
            match_term = clean_num.replace(" ", "")
            name_term = re.sub(r'[^가-힣0-9]', '', str(task.get("name", "")))[:10] # 한글/숫자만 10자
            # 국방부 공고 감지: '국방부' 텍스트 OR 국방부 특유 번호 패턴(LNG/UN/MND 등)
            _num_str = str(task.get("num", ""))
            is_mnd = ("국방부" in _num_str) or bool(re.match(r'^(LNG|UN|MND|UE|UF|UG|UH|UI|UJ)', _num_str, re.IGNORECASE))
            
            for title_elem in titles:
                try:
                    title_text = title_elem.text.strip()
                    if "입찰" in title_text and "결과" not in title_text:
                        continue
                        
                    print(f"   [디버그] '{title_text}' 섹션 매칭 시도 (국방부:{is_mnd})")
                    
                    # 결과공고 섹션의 모든 행(tr) 수집
                    # ancestor::div[1] 에서 sibling div를 5개까지 탐색 (구매/공사/용역/매각 탭 대응)
                    rows = []
                    try:
                        parent_containers = title_elem.find_elements(
                            By.XPATH, "./ancestor::div[1]/following-sibling::div[position() <= 5]"
                        )
                        for container in parent_containers:
                            rows.extend(container.find_elements(By.TAG_NAME, "tr"))
                    except:
                        pass
                    
                    if not rows:
                        try:
                            rows = title_elem.find_elements(By.XPATH, "./following::tr[position() <= 20]")
                        except:
                            pass

                    for idx, row in enumerate(rows):
                        try:
                            links = row.find_elements(By.TAG_NAME, "a")
                            if not links:
                                continue  # 헤더 행 건너뜀

                            row_text = (row.get_attribute("innerText") or row.text).replace(" ", "").replace("\n", "")
                            row_text_ko = re.sub(r'[^가-힣0-9]', '', row_text)

                            # [국방부 특수 처리] 공고번호가 검색어↔상세페이지 간 불일치하므로
                            # 결과공고 영역에 공고가 1개 이상 있으면 → 첫 번째 링크 행을 바로 선택
                            if is_mnd:
                                target_link = links[0]
                                print(f"   ✅ [국방부] 결과공고 첫 번째 항목 자동 선택: {row_text[:40]}")
                                break

                            # 일반 공고: 공고번호 매칭
                            if match_term and match_term in row_text:
                                target_link = links[0]
                                print(f"   ✅ 번호 매칭 성공: {match_term}")
                                break

                            # 일반 공고: 공고명 매칭 (번호 불일치 대비)
                            if name_term and name_term in row_text_ko:
                                target_link = links[0]
                                print(f"   ✅ 공고명 유사 매칭 성공: {name_term}")
                                break

                        except: continue
                        if target_link: break
                    if target_link: break
                except: continue

            
            # 2. 전수 조사 (최후의 수단: 페이지 전체에서 '결' 배지가 있는 행 탐색)
            if not target_link:
                print("   [디버그] 섹션 기반 탐색 실패, 페이지 전체 전수 조사 시작...")
                first_result_link = None  # 결과공고 섹션 첫 번째 링크(폴백용)
                # 모든 tr을 가져와서 '결' 아이콘과 번호가 동시에 있는 행 찾기
                all_rows = self.driver.find_elements(By.TAG_NAME, "tr")
                for row in all_rows:
                    try:
                        row_html = row.get_attribute("innerHTML")
                        if 'alt="결"' in row_html:
                            row_links = row.find_elements(By.TAG_NAME, "a")
                            if row_links and first_result_link is None:
                                first_result_link = row_links[0]  # 폴백: 결 배지 있는 첫 번째 행
                            row_text = (row.get_attribute("innerText") or row.text).replace(" ", "")
                            if match_term in row_text:
                                target_link = row_links[0] if row_links else None
                                print(f"   ✅ 페이지 전수 조사(결 배지)로 결과 항목 발견")
                                break
                    except: continue
                
                # 3. 폴백: 번호 매칭 실패해도 결과공고 섹션에 항목이 1개뿐이면 자동 선택
                if not target_link and first_result_link:
                    target_link = first_result_link
                    print(f"   ✅ [폴백] 결과공고 섹션 첫 번째 항목 자동 선택 (번호 매칭 불일치 대응)")

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
            os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
            with open(f"{KbidConfig.DEBUG_HISTORY_FOLDER}/detail_page_debug.html", "w", encoding="utf-8") as f:
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
        """결과 데이터 추출 (공고기관, 참여업체, 사정률, 1등, AIR/에어 등)"""
        data = {
            "공고기관": self.get_agency(),
            "참여 업체수": "", "사정률": "", "1등 상호명": "",
            "1등 업체 입찰금액": "", "1등 업체 사정률": "",
            "AIR 채호원 입찰금액": "-", "AIR 채호원 사정률": "-", "AIR 채호원 순위": "-",
            "에어채호원 입찰금액": "-", "에어채호원 사정률": "-", "에어채호원 순위": "-",
            "A값": ""
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
            
            # [디버그] 채호원 검색 후 페이지 HTML 저장
            try:
                os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
                with open(f"{KbidConfig.DEBUG_HISTORY_FOLDER}/chaehowon_search_debug.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                print("   [디버그] 채호원 검색 결과 HTML 저장: chaehowon_search_debug.html")
            except Exception as e_html:
                print(f"   [디버그] HTML 저장 실패: {e_html}")
            
            headers, rows = self._get_table_data()
            print(f"   [디버그] 테이블 헤더: {headers}")
            print(f"   [디버그] 테이블 행 수: {len(rows)}")
            for i, row in enumerate(rows):
                print(f"   [디버그] 행[{i}]: {row}")
            
            found_air = False
            found_corp = False
            
            for row in rows:
                # '상호명' 또는 '상호' 컬럼 모두 대응
                name = (self._get_cell(row, headers, "상호명") or self._get_cell(row, headers, "상호")).replace(" ", "")
                # '업체사정률' 또는 '사정률' 컬럼 모두 대응
                def get_rate(r, h):
                    return self._get_cell(r, h, "업체사정률") or self._get_cell(r, h, "사정률")
                
                if "AIR채호원" in name and not found_air:
                    data["AIR 채호원 입찰금액"] = self._format_amount(self._get_cell(row, headers, "입찰금액"))
                    data["AIR 채호원 사정률"] = self._format_rate(get_rate(row, headers))
                    data["AIR 채호원 순위"] = self._get_cell(row, headers, "순위")
                    found_air = True
                    print(f"   ✅ AIR 채호원 발견: 순위={data['AIR 채호원 순위']}, 사정률={data['AIR 채호원 사정률']}")
                if ("에어채호원" in name or "애어체호원" in name or "에어체호원" in name) and not found_corp:
                    data["에어채호원 입찰금액"] = self._format_amount(self._get_cell(row, headers, "입찰금액"))
                    data["에어채호원 사정률"] = self._format_rate(get_rate(row, headers))
                    data["에어채호원 순위"] = self._get_cell(row, headers, "순위")
                    found_corp = True
                    print(f"   ✅ 에어채호원 발견: 순위={data['에어채호원 순위']}, 사정률={data['에어채호원 사정률']}")
                if found_air and found_corp: break
            
            if found_air or found_corp:
                print(f"   ✅ 대상 업체 발견: AIR({data['AIR 채호원 순위']}위), 에어({data['에어채호원 순위']}위)")
            else:
                print("   ⚠️ 검색 결과 내에 대상 업체(채호원)가 없습니다.")
        except Exception as e:
            print(f"   ⚠️ 채호원 검색 중 오류: {e}")
            import traceback; traceback.print_exc()

    def _find_text_by_label(self, label):
        try:
            xpath = f"//th[contains(., '{label}')]/following-sibling::td[1]"
            return self.driver.find_element(By.XPATH, xpath).text.strip()
        except: return ""

    def get_agency(self):
        """공고기관 정보 추출 (img 태그 제거)"""
        try:
            xpath = "//th[contains(text(), '공고기관')]/following-sibling::td[1]"
            element = self.driver.find_element(By.XPATH, xpath)
            # JavaScript로 img 태그 제거 후 텍스트만 추출
            script = """
                var elem = arguments[0];
                var clone = elem.cloneNode(true);
                var imgs = clone.querySelectorAll('img');
                imgs.forEach(img => img.remove());
                return clone.textContent.trim();
            """
            text = self.driver.execute_script(script, element)
            return text.strip() if text else ""
        except:
            return ""


    def _get_table_data(self):
        """다양한 셀렉터로 테이블을 시도하고 데이터 반환"""
        table = None
        # 낙찰순위 전용 XPath를 우선순위 최상위로 추가 (개찰결과 테이블과 구분)
        extra_selectors = [
            "//h4[contains(.,'낙찰순위')]/following-sibling::table[1]",  # "낙찰순위" 제목 바로 다음 형제 테이블
            "//div[contains(@class,'result_com_search')]/following-sibling::table[1]",  # 검색 폼 다음 테이블
            "//table[@class='tbl_result_coms' and not(contains(@class, 'tbl_result_one'))]",  # 결과 테이블 중 '개찰결과' 제외
            "//h3[contains(.,'낙찰순위') or contains(.,'개찰결과')]/following::table[1]",
            "//div[contains(@class,'result') or contains(@id,'Result')]//table",
            "//caption[contains(.,'낙찰순위')]/ancestor::table",
        ]
        selectors = extra_selectors + list(KbidConfig.SELECTORS["ranking_table"])
        
        for sel in selectors:
            try:
                if sel.startswith("//") or sel.startswith("("):
                    table = self.driver.find_element(By.XPATH, sel)
                elif sel.startswith("#"):
                    table = self.driver.find_element(By.ID, sel[1:])
                else:
                    table = self.driver.find_element(By.CSS_SELECTOR, sel)
                if table and table.is_displayed():
                    print(f"   [디버그] 테이블 셀렉터 매칭: {sel[:60]}")
                    break
                table = None
            except: continue
        
        # 마지막 폴백: 페이지의 모든 table 중 td가 가장 많은 것
        if not table:
            all_tables = self.driver.find_elements(By.TAG_NAME, "table")
            best, best_count = None, 0
            for t in all_tables:
                try:
                    cnt = len(t.find_elements(By.TAG_NAME, "td"))
                    if cnt > best_count:
                        best_count = cnt
                        best = t
                except: continue
            if best_count > 0:
                table = best
                print(f"   [디버그] 폴백: td가 가장 많은 테이블 선택 (td={best_count}개)")

        if not table:
            print("   [디버그] 테이블을 찾지 못했습니다.")
            return [], []

        try:
            header_elements = table.find_elements(By.XPATH, ".//th")
            headers = [h.text.strip() for h in header_elements]
            
            rows = []
            tr_elements = table.find_elements(By.XPATH, ".//tr[td]")
            for tr in tr_elements:
                cells = [c.text.strip() for c in tr.find_elements(By.TAG_NAME, "td")]
                # 헤더보다 셀 수가 적어도 데이터가 있으면 허용 (colspan 등 대응)
                if cells:
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
        print("🔄 [1/3] 구글 시트 연결 중...")
        self.gs = GoogleSheetsManager()
        print("✅ 구글 시트 연결 완료")
        
        print("🔄 [2/3] 크롬 드라이버 초기화 중...")
        self.browser = KbidBrowser()
        print("✅ 크롬 드라이버 초기화 완료")
        
        print("🔄 [3/3] 파서 초기화 중...")
        self.parser = KbidParser(self.browser.driver)
        print("✅ 파서 초기화 완료")

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
    os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
    KbidResultCrawler().run()
