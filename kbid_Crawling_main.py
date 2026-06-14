# -*- coding: utf-8 -*-
"""
KBID 입찰 정보 크롤러 (구글 시트 연동 버전)
파일명: kbid_Crawling_main.py

[사전 준비]
1. pip install gspread google-auth selenium pandas
2. credentials.json 파일(구글 API 인증키)을 같은 폴더에 준비
3. 구글 시트 이름: '입찰관리' (시트 내에 '입찰공고', '투찰준비' 탭 필요)
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
    SEARCH_URL_TEMPLATE = "https://www.kbid.co.kr/search/index.htm?mid=lge123&txtFindWordTop={}"
    
    # 구글 시트 설정 (OAuth 2.0 방식으로 변경)
    CLIENT_SECRETS_FILE = "client_secrets.json"
    TOKEN_FILE = "token.json"  # 로그인 후 자동 생성됨
    SPREADSHEET_NAME = "입찰관리"
    DEBUG_HISTORY_FOLDER = "debug_history"
    
    SELECTORS = {
        "login_check": [
            "//*[contains(text(), '로그아웃')]",
            "//*[contains(text(), '로그 아웃')]",
            "//*[contains(text(), '마이페이지')]",
            "//a[contains(@href, 'logout')]"
        ],
        "search_results": "#listBody1 tr, tbody#listBody1 tr, table#listBody1 tr, .search-result tr, tbody tr, table tr",
        "bid_links": "#listBody1 a[href]",
        "detail_title": "td.h_tit",
        "schedule_date": "//div[contains(@class, 'scheduler_flowchart')]//li[contains(@class, 'on')]//div[@class='area_date']",
        "th_td_pair": "(//th[contains(text(), '{label}')])[last()]/following-sibling::td[1]",
        "bid_agency": "(//th[contains(text(), '공고기관')])[last()]/following-sibling::td[1]"
    }

class GoogleSheetsManager:
    """구글 시트 데이터 입출력 관리 (OAuth 2.0 적용)"""
    def __init__(self):
        if not os.path.exists(KbidConfig.CLIENT_SECRETS_FILE):
            raise FileNotFoundError(f"'{KbidConfig.CLIENT_SECRETS_FILE}' 파일이 없습니다. OAuth JSON 파일을 준비해주세요.")
            
        # gspread.oauth()는 브라우저를 열어 사용자 인증을 처리합니다.
        self.client = gspread.oauth(
            credentials_filename=KbidConfig.CLIENT_SECRETS_FILE,
            authorized_user_filename=KbidConfig.TOKEN_FILE
        )
        
        # 시트가 없으면 새로 생성하는 로직 추가
        try:
            self.sheet = self.client.open(KbidConfig.SPREADSHEET_NAME)
            self._ensure_prepare_headers()
        except gspread.exceptions.SpreadsheetNotFound:
            print(f"✨ '{KbidConfig.SPREADSHEET_NAME}' 시트를 찾을 수 없어 새로 생성합니다...")
            self.sheet = self.client.create(KbidConfig.SPREADSHEET_NAME)
            self._initialize_sheets()
            print("✅ 시트 구조 생성이 완료되었습니다. 구글 드라이브에서 '입찰관리' 시트를 열어")
            print("   '입찰공고' 탭의 '*공고명' 아래에 작업할 공고명을 입력 후 다시 실행해 주세요.")
            os._exit(0) # 처음 생성 시에는 데이터가 없으므로 즉시 종료

    def _initialize_sheets(self):
        """기본 시트 구성 및 헤더 세팅"""
        # 1. 입찰공고 시트 (읽기용) 생성
        try:
            ws_announce = self.sheet.add_worksheet(title="입찰공고", rows="100", cols="10")
            ws_announce.append_row(["*공고명", "공고번호", "기타메모"]) 
        except: pass
        
        # 2. 투찰준비 시트 (쓰기용) 생성
        try:
            ws_prepare = self.sheet.add_worksheet(title="투찰준비", rows="100", cols="25")
            headers = [
                "투찰상태", "공고명", "공고번호", "공고기관", "지역제한", "업종", "입찰개시일", "투찰마감일시", "개찰일시",
                "기초금액", "A값", "예가변동폭", "투찰하한율", "계약방법",
                "예상투찰가1", "예상투찰가2", "예상투찰가3",
                "참여 업체수", "사정률", "1등 상호명", "1등 업체 입찰금액", "1등 업체 사정률",
                "AIR 채호원 입찰금액", "AIR 채호원 사정률", "AIR 채호원 순위",
                "에어채호원 입찰금액", "에어채호원 사정률", "에어채호원 순위"
            ]
            ws_prepare.append_row(headers)
        except: pass
        
        # 기본 'Sheet1' 삭제
        try:
            default_ws = self.sheet.get_worksheet(0)
            if default_ws.title == "Sheet1":
                self.sheet.del_worksheet(default_ws)
        except: pass

    def _ensure_prepare_headers(self):
        """기존 '투찰준비' 시트 헤더를 최신 구조로 재구성"""
        try:
            ws = self.sheet.worksheet("투찰준비")
        except gspread.exceptions.WorksheetNotFound:
            return

        desired_headers = [
            "투찰상태", "공고명", "공고번호", "공고기관", "지역제한", "업종", "입찰개시일", "투찰마감일시", "개찰일시",
            "기초금액", "A값", "예가변동폭", "투찰하한율", "계약방법",
            "예상투찰가1", "예상투찰가2", "예상투찰가3",
            "참여 업체수", "사정률", "1등 상호명", "1등 업체 입찰금액", "1등 업체 사정률",
            "AIR 채호원 입찰금액", "AIR 채호원 사정률", "AIR 채호원 순위",
            "에어채호원 입찰금액", "에어채호원 사정률", "에어채호원 순위"
        ]

        current_headers = [h.replace("*", "").strip() for h in ws.row_values(1)]
        if current_headers == desired_headers:
            return

        all_values = ws.get_all_values()
        new_rows = [desired_headers]

        for row in all_values[1:]:
            new_row = ["" for _ in desired_headers]
            for idx, val in enumerate(row):
                if idx >= len(current_headers):
                    continue
                header_name = current_headers[idx]
                if header_name in desired_headers:
                    new_row[desired_headers.index(header_name)] = val
            new_rows.append(new_row)

        ws.clear()
        ws.update("A1", new_rows)

    def get_search_terms(self):
        """'입찰공고' 시트에서 '공고명' 및 '공고번호' 데이터를 가져옵니다."""
        ws = self.sheet.worksheet("입찰공고")
        all_data = ws.get_all_records()
        
        tasks = []
        for row in all_data:
            # 헤더 이름에 '*'가 포함될 수 있으므로 유연하게 처리
            name = next((v for k, v in row.items() if "공고명" in k), "")
            num = next((v for k, v in row.items() if "공고번호" in k), "")
            if name or num:
                # 상태 표시 접두사 제거 (반드시 공백이 있는 경우에만 제거하여 '전기', '전자' 등이 '기', '자'로 훼손되는 것 방지)
                clean_name = re.sub(r'^(?:결|전|수|취)\s+', '', str(name)).strip()
                clean_name = clean_name.replace("(과거 공고)", "").replace("(취소)", "").strip()
                tasks.append({"name": clean_name, "num": str(num).strip()})
        return tasks

    def get_processed_bids(self):
        """이미 '투찰준비' 시트에 기록된 공고번호 목록을 가져옵니다."""
        try:
            ws = self.sheet.worksheet("투찰준비")
            # 두 번째 컬럼(공고번호)의 모든 값을 가져옴 (B열)
            return set(ws.col_values(2)[1:])
        except:
            return set()

    def _parse_datetime(self, text):
        if not text:
            return None

        text = text.replace('\xa0', ' ')
        text = text.replace('년', '-').replace('월', '-').replace('일', ' ').replace('시', ':').replace('분', ' ')
        text = re.sub(r'[^0-9:\-\.\/ ]', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip()

        patterns = [
            r'\d{4}[./\-]\d{1,2}[./\-]\d{1,2}\s+\d{1,2}:\d{1,2}:?\d{0,2}',
            r'\d{4}[./\-]\d{1,2}[./\-]\d{1,2}\s+\d{1,2}:\d{1,2}',
            r'\d{4}[./\-]\d{1,2}[./\-]\d{1,2}',
        ]

        for pat in patterns:
            match = re.search(pat, text)
            if match:
                candidate = match.group().strip()
                candidate = candidate.replace('.', '-').replace('/', '-')
                parts = re.findall(r'\d+', candidate)
                if len(parts) >= 5:
                    year, month, day, hour, minute = map(int, parts[:5])
                    return datetime(year, month, day, hour, minute)
                if len(parts) >= 3:
                    year, month, day = map(int, parts[:3])
                    return datetime(year, month, day, 0, 0)
        return None

    def update_bid_statuses(self):
        """투찰준비 시트의 개찰일시를 기준으로 상태를 갱신합니다."""
        try:
            ws = self.sheet.worksheet("투찰준비")
        except gspread.exceptions.WorksheetNotFound:
            return

        headers = [h.replace("*", "").strip() for h in ws.row_values(1)]
        try:
            status_idx = headers.index("투찰상태")
            open_idx = headers.index("개찰일시")
            close_idx = headers.index("투찰마감일시") if "투찰마감일시" in headers else None
        except ValueError:
            print("⚠️ 투찰준비 시트의 필수 컬럼을 찾을 수 없습니다.")
            print(f"   현재 헤더: {headers}")
            return

        rows = ws.get_all_values()[1:]
        updates = []
        now = datetime.now()
        
        print(f"📊 투찰준비 시트 검사: 총 {len(rows)}행, 현재시간 {now.strftime('%Y-%m-%d %H:%M')}")
        print(f"   헤더 인덱스 - 투찰상태({status_idx}), 개찰일시({open_idx}), 투찰마감({close_idx})")

        for row_idx, row in enumerate(rows, start=2):
            # 행이 비어있으면 건너뛰기
            if not row or all(cell.strip() == "" for cell in row):
                continue
                
            current_status = row[status_idx].strip() if len(row) > status_idx else ""
            bid_no = row[1].strip() if len(row) > 1 else ""
            
            # 결과확인 상태는 더 이상 변경하지 않음
            if current_status == "결과확인":
                continue

            open_text = row[open_idx].strip() if len(row) > open_idx else ""
            open_time = self._parse_datetime(open_text)
            close_time = None
            close_text = ""
            if close_idx is not None and len(row) > close_idx:
                close_text = row[close_idx].strip()
                close_time = self._parse_datetime(close_text)

            print(f"  [{row_idx}] {bid_no}: 상태='{current_status}' | 개찰={open_text} (파싱: {open_time}) | 투찰마감={close_text if close_idx else 'N/A'}")

            # 결정할 새로운 상태
            new_status = None
            
            # 1. 상태가 비어있으면 기본값으로 "투찰대기" 설정
            # (추가: 실제로 데이터가 있으면 아래 로직으로 자동 결정)
            if not current_status:
                # 투찰마감일시 기준으로 상태 자동 결정
                if close_time and now >= close_time:
                    new_status = "투찰완료"  # 마감 지남 (개찰 전)
                elif open_time and now >= open_time:
                    new_status = "결과확인"  # 개찰 지남
                else:
                    new_status = "투찰대기"  # 투찰 가능 기간
                print(f"       → 상태 비어있음: 자동으로 '{new_status}'로 설정")
            
            # 2. 이미 상태가 있으면 개찰일시 기준으로만 "결과확인"으로 업데이트
            elif current_status == "투찰대기" or current_status == "투찰완료":
                if open_time and now >= open_time:
                    new_status = "결과확인"
                    print(f"       → 개찰일시 경과: {open_time.strftime('%Y-%m-%d %H:%M')} <= {now.strftime('%Y-%m-%d %H:%M')}")
                elif close_time and now >= close_time and not open_time:
                    # 개찰일시가 없고 투찰마감만 지난 경우
                    new_status = "투찰완료"
                    print(f"       → 투찰마감 경과: {close_time.strftime('%Y-%m-%d %H:%M')} <= {now.strftime('%Y-%m-%d %H:%M')}")

            # 3. 상태 변경 필요 시 업데이트 목록에 추가
            if new_status:
                cell_name = gspread.utils.rowcol_to_a1(row_idx, status_idx + 1)
                updates.append((cell_name, new_status))

        if updates:
            batch_data = [{"range": cell, "values": [[value]]} for cell, value in updates]
            ws.batch_update(batch_data)
            print(f"✅ {len(updates)}건의 투찰상태를 업데이트했습니다.")
            return len(updates)
        else:
            print("  → 상태 변경 필요 항목 없음")
        return 0

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

    def _find_duplicate_bid_row(self, bid_no, ws):
        """같은 기본 공고번호를 가진 기존 행을 찾아 비교
        새 공고가 더 최신이면 기존 행을 삭제하고 None 반환
        같은 공고면 그 행 인덱스 반환
        """
        all_bid_nos = ws.col_values(2)[1:]  # 헤더 제외
        
        new_base, new_serial = self._extract_bid_base_and_serial(bid_no)
        
        for idx, sheet_bid in enumerate(all_bid_nos):
            sheet_base, sheet_serial = self._extract_bid_base_and_serial(sheet_bid)
            
            # 기본 부분이 같은 경우
            if new_base and sheet_base and new_base.upper() == sheet_base.upper():
                # 새 일련번호가 더 크면: 기존 행 삭제
                if new_serial > sheet_serial:
                    row_idx_to_delete = idx + 2  # 1-based index (헤더는 1, 데이터는 2부터)
                    ws.delete_rows(row_idx_to_delete)
                    print(f"🗑️  재공고 감지: 기존 공고 삭제 ({sheet_bid}) -> 새 공고 ({bid_no})")
                    return None
                # 새 일련번호가 같거나 작으면: 기존 행 반환
                elif new_serial <= sheet_serial:
                    return idx + 2  # 1-based index
        
        return None

    def save_result(self, data_dict):
        """추출된 데이터를 '투찰준비' 시트에 저장 (중복 시 업데이트, 재공고 시 최신만 기록)"""
        try:
            ws = self.sheet.worksheet("투찰준비")
            headers = ws.row_values(1)
            bid_no = data_dict.get("공고번호", "").strip()
            
            # 공고명에서 상태 표시 접두사 제거 (결, 전, 수, 취 등) 및 '(과거 공고)' 제거
            if "공고명" in data_dict:
                original_name = data_dict["공고명"]
                # 상태 표시 접두사 제거 (반드시 공백이 있는 경우에만 제거하여 '전주대' 등이 '주대'로 훼손되는 것 방지)
                cleaned_name = re.sub(r"^(결|전|수|취|긴|견|재)\s+", "", original_name).strip()
                # (과거 공고) 제거
                cleaned_name = cleaned_name.replace("(과거 공고)", "").strip()
                data_dict["공고명"] = cleaned_name

            # 만약 공고명에 '(취소)'가 포함되어 있으면 투찰상태를 '공고 취소'로 설정
            if "공고명" in data_dict and "(취소)" in data_dict["공고명"]:
                data_dict["투찰상태"] = "공고 취소"
                data_dict["공고명"] = data_dict["공고명"].replace("(취소)", "").strip() # 공고명에서 (취소) 제거

            # 기존 데이터에서 공고번호 열(2번째 열) 검색 및 중복/재공고 처리
            row_idx = self._find_duplicate_bid_row(bid_no, ws)
                    
            if row_idx is not None:
                # 이미 존재하는 경우 해당 행 번호 찾기 (1-based index)
                
                # 기존 데이터 가져오기 (API 호출 최소화를 위해 범위 지정)
                end_col = chr(64 + len(headers)) if len(headers) <= 26 else "Z"
                range_name = f"A{row_idx}:{end_col}{row_idx}"
                existing_row_data = ws.get(range_name)
                existing_row = existing_row_data[0] if existing_row_data else []
                existing_row = existing_row + [""] * (len(headers) - len(existing_row))
                
                row_to_save = []
                for i, h in enumerate(headers):
                    clean_h = h.replace("*", "").strip()
                    if clean_h in data_dict:
                        row_to_save.append(data_dict[clean_h])
                    else:
                        row_to_save.append(existing_row[i])
                
                # 해당 행 업데이트
                ws.update(range_name, [row_to_save])
                print(f"🔄 기존 데이터 업데이트 완료: {bid_no}")
            else:
                # 존재하지 않는 경우 새로 추가
                row_to_save = []
                for h in headers:
                    clean_h = h.replace("*", "").strip()
                    row_to_save.append(data_dict.get(clean_h, ""))
                    
                ws.append_row(row_to_save)
                print(f"✅ 새 데이터 추가 완료: {bid_no} (공고번호에 일련번호까지 표시)")
            
            return True
        except Exception as e:
            print(f"❌ 구글 시트 저장 오류: {e}")
            traceback.print_exc()
            return False

class KbidBrowser:
    """브라우저 자동화 및 탐지 우회 제어"""
    def __init__(self, gs_manager):
        self.gs = gs_manager # GoogleSheetsManager 인스턴스 저장
        self.driver = self._init_driver()

    def _init_driver(self):
        options = Options()
        options.add_argument("--incognito")
        options.add_argument("--log-level=3")  # 크롬 내부 로그(GCM 등) 억제
        options.add_argument("--disable-logging")
        options.page_load_strategy = 'normal'  # 'eager' 대신 'normal' 사용하되 타임아웃 관리
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0')
        
        driver = webdriver.Chrome(options=options)
        # 탐지 우회를 위한 스크립트 실행
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        # 암묵적 대기: 2초로 설정 (동적 요소 렌더링 대기)
        driver.set_page_load_timeout(30)
        driver.implicitly_wait(2)
        return driver

    def wait_for_login(self):
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
                break
            elif user_input == 'n':
                print("⏳ 로그인을 완료한 후 다시 입력해주세요.")
                continue
            else:
                print("⚠️ y 또는 n을 입력해주세요.")
                continue
        
        # 검색 페이지로 이동
        print("   [공고 검색 페이지 이동] 진행 중...")
        try:
            self.driver.get(KbidConfig.SEARCH_URL_TEMPLATE.format(""))
            
            # 페이지 로드 완료 대기 (최대 15초)
            for attempt in range(15):
                try:
                    page_state = self.driver.execute_script("return document.readyState")
                    if page_state == "complete":
                        print("   ✅ 공고 검색 페이지 로드 완료\n")
                        return True
                except:
                    pass
                time.sleep(1)
            
            # 15초 대기 후 성공으로 간주
            print("   ✅ 공고 검색 페이지 로드 완료 (타임아웃)\n")
            return True
            
        except Exception as e:
            print(f"   ⚠️ 공고 검색 페이지 이동: {str(e)[:100]}")
            print("   계속 진행합니다...")
            return True

    def navigate_to_bid(self, task):
        """공고명/번호로 검색 후 정확한 공고번호 확인하여 상세 페이지 이동"""
        name = task.get("name", "")
        raw_num = task.get("num", "")
        
        # 핵심 번호만 추출 (예: R26BK01488658-000, 202604021057-00)
        # 문자열 내에서 영문+숫자+하이픈 조합의 패턴을 찾음
        match = re.search(r'[A-Z0-9]{4,}-[A-Z0-9]+', raw_num)
        clean_num = match.group() if match else raw_num.strip()
        
        # 검색어 결정: 공고번호를 우선 사용 (공고명 검색 실패 시)
        search_term_raw = clean_num if clean_num else name
        if not search_term_raw: return False

        # 공고명일 경우 괄호 접두사 제거
        if not clean_num:
            search_term_raw = re.sub(r'^[\(\[\{]\s*[^\)\]\}]+\s*[\)\]\}]\s*', '', search_term_raw).strip()
        
        search_term = quote_plus(search_term_raw)
        # URL 결정 및 로그 출력
        url = KbidConfig.SEARCH_URL_TEMPLATE.format(search_term)
        
        # [최적화] 메인페이지 정체 방지 및 강력한 탈출 로직
        for attempt in range(3):
            try:
                # 0. 윈도우 상태 관리 (불필요한 팝업창 닫기)
                if len(self.driver.window_handles) > 1:
                    main_handle = self.driver.window_handles[0]
                    for handle in self.driver.window_handles[1:]:
                        if handle != main_handle:
                            self.driver.switch_to.window(handle)
                            self.driver.close()
                    self.driver.switch_to.window(main_handle)

                # 1. 현재 상태 확인
                try:
                    decoded_current_url = unquote(self.driver.current_url)
                except:
                    time.sleep(2)
                    continue

                if search_term_raw in decoded_current_url and "search/index.htm" in decoded_current_url:
                    print("   [디버그] 이미 해당 검색 결과 페이지에 있습니다. 이동을 생략합니다.")
                    break
                
                print(f"   [디버그] 페이지 이동 시도 ({attempt+1}/3): {url}")
                
                # 2. 이동 방식 결정: 메인 페이지라면 직접 입력을 우선 시도 (URL 이동이 막히는 경우 대비)
                if "index_first" in decoded_current_url or decoded_current_url.endswith(".co.kr/"):
                    search_input = self.driver.find_elements(By.ID, "s_search_word")
                    if search_input and search_input[0].is_displayed():
                        print("   [디버그] 메인페이지 검색창 직접 입력 시도")
                        search_input[0].clear()
                        search_input[0].send_keys(search_term_raw)
                        search_input[0].send_keys("\n")
                    else:
                        self.driver.get(url)
                else:
                    self.driver.set_page_load_timeout(20)
                    self.driver.get(url)

                time.sleep(2.0)
                
                # 3. 검색 결과 페이지의 핵심 요소 확인
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='idCBidTable'] | //*[contains(@class, 'search_list')] | //*[@id='listBody1']"))
                )
                print("   ✅ 검색 결과 페이지 진입 성공")
                break
            except Exception as e:
                print(f"   ⚠️ 검색 페이지 진입 지연 ({attempt+1}/3): {str(e)[:100]}")
                try:
                    self.driver.execute_script("window.stop();")
                    self.driver.switch_to.alert.accept()
                except: pass
                time.sleep(2)
                if attempt == 2:
                    print(f"   ⚠️ 최종 페이지 로드 지연 (무시하고 진행): {e}")
        
        # 로그인 페이지로 튕겼는지 확인 (URL 또는 페이지 내용 확인)
        current_url = self.driver.current_url.lower()
        is_login_page = "login" in current_url
        
        # URL에 login이 없어도 로그인 폼(MemID 필드)이 보이면 로그인 페이지로 간주 (KBID 마스킹 대응)
        if not is_login_page:
            try:
                mem_id_elements = self.driver.find_elements(By.ID, "MemID")
                if mem_id_elements and mem_id_elements[0].is_displayed():
                    is_login_page = True
            except: pass
            
        if is_login_page:
            print("   ⚠️ 세션 만료 또는 로그아웃 감지 - 재로그인 대기")
            if not self.wait_for_login(): return False
            
            # 로그인 성공 후 바로 검색어로 이동 시도
            try:
                # 1안: 검색창에 직접 입력 (wait_for_login이 이미 검색 페이지에 있을 것이므로 바로 입력 시도)
                time.sleep(1)
                search_inputs = self.driver.find_elements(By.ID, "s_search_word")
                if search_inputs:
                    search_input = search_inputs[0]
                    search_input.clear()
                    search_input.send_keys(search_term_raw)
                    time.sleep(0.3)
                    # 검색 버튼 클릭 또는 엔터
                    search_btn_elements = self.driver.find_elements(By.ID, "search_btn")
                    if search_btn_elements:
                        self.driver.execute_script("arguments[0].click();", search_btn_elements[0])
                    else:
                        search_input.send_keys("\n")
                    print(f"   [디버그] 로그인 확인 후 직접 입력 방식으로 검색 진행")
                else:
                    # 검색창이 안 보이면 URL 직접 이동 (타임아웃 적용)
                    self.driver.set_page_load_timeout(15)
                    self.driver.get(url)
            except:
                try:
                    self.driver.set_page_load_timeout(15)
                    self.driver.get(url)
                except:
                    try: self.driver.execute_script("window.stop();")
                    except: pass
            time.sleep(1.5)

        try:
            print(f"   [디버그] 현재 창 개수: {len(self.driver.window_handles)}", flush=True)
            print(f"   [디버그] 현재 URL: {self.driver.current_url}", flush=True)
            
            # 모든 창 확인
            for i, handle in enumerate(self.driver.window_handles):
                try:
                    self.driver.switch_to.window(handle)
                    print(f"      - 창 {i} URL: {self.driver.current_url}")
                except: pass
            
            # 원래 창으로 복귀 (마지막 창 기준)
            self.driver.switch_to.window(self.driver.window_handles[-1])

            # 디버깅을 위해 페이지 소스 저장
            try:
                os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
                with open(f"{KbidConfig.DEBUG_HISTORY_FOLDER}/search_debug.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                print("   [디버그] search_debug.html 저장 완료")
            except: pass
            
            if "잘못된 접근" in self.driver.page_source:
                print("   [디버그] '잘못된 접근' 감지 - 재접속 시도")
                self.driver.get("https://www.kbid.co.kr/search/index.htm?mid=lge123")
                time.sleep(0.5)
                self.driver.get(url)
                time.sleep(1)

            # 탐색 함수: 메인 프레임 우선 탐색 및 선택적 프레임 탐색
            # (kbid.co.kr 검색 결과는 항상 메인 프레임에 있으며,
            #  광고/추적 iframe 순회가 120초 타임아웃을 유발했던 핵심 원인)
            def find_in_frames(selector, search_frames=False):
                try:
                    by_type = By.XPATH if selector.startswith("/") or selector.startswith("(") else By.CSS_SELECTOR
                    
                    # 1. 메인 프레임 확인 (KBID 검색 결과는 대부분 메인 프레임에 있음)
                    self.driver.switch_to.default_content()
                    elements = self.driver.find_elements(by_type, selector)
                    if elements: return elements
                    
                    # 프레임 탐색이 필요 없는 경우 즉시 종료 (검색 결과 테이블 등)
                    if not search_frames:
                        return []

                    # 2. 아이프레임 및 프레임 탐색 (최소한으로 제한하여 지연 방지)
                    all_frames = self.driver.find_elements(By.XPATH, "//iframe | //frame")
                    if len(all_frames) > 3: # 5개에서 3개로 더 축소
                        all_frames = all_frames[:3]
                        
                    for f in all_frames:
                        try:
                            self.driver.switch_to.frame(f)
                            elements = self.driver.find_elements(by_type, selector)
                            if elements: return elements
                            self.driver.switch_to.default_content()
                        except:
                            try: self.driver.switch_to.default_content()
                            except: pass
                            continue
                except Exception:
                    pass
                
                try: self.driver.switch_to.default_content()
                except: pass
                return []

            # 페이지 로드 후 아이프레임/AJAX 로딩 대기
            time.sleep(2.0)
            
            # 검색 단계에서는 즉시 응답하도록 암묵적 대기 0초 설정
            self.driver.implicitly_wait(0)
            
            # 1단계: 페이지 로드 및 AJAX 테이블 확인
            print(f"⏳ 페이지 로드 및 데이터 수집 대기 중 (키워드: {search_term_raw[:20]}...)...")
            
            table_found = False
            for i in range(7): # 최대 7초간 테이블 컨테이너 확인
                # KBID 검색 결과 테이블 ID: idCBidTable
                if find_in_frames("#idCBidTable, .tbl_search_list, #listBody1"):
                    table_found = True
                    break
                
                # 중간에 로그인 페이지로 변했는지 재확인 (page_source 대신 current_url 사용으로 성능 최적화)
                if i == 3:
                    try:
                        if "login" in self.driver.current_url.lower():
                            print("   ⚠️ 로딩 중 로그인 페이지 감지 - 재로그인 대기")
                            if not self.wait_for_login(): return False
                            self.driver.get(url)
                    except: pass
                
                time.sleep(1)
            
            if not table_found:
                print("   ⚠️ 결과 테이블 컨테이너를 찾지 못했습니다. (프레임 확인 중...)")

            # 2단계: 비동기 데이터(AJAX) 로딩 대기 및 행 탐색
            print("⏳ 입찰공고 데이터 로딩 대기 중...")
            bid_rows = []
            
            # AJAX로 tbody가 채워지는 것을 기다림 (최대 10초)
            for wait_step in range(10):
                xpath_candidates = [
                    "//tbody[@id='listBody1']//tr[td//a]", # 최근 입찰공고 전용 리스트
                    "//tr[contains(@onclick, 'Gview') and ancestor::tbody[@id='listBody1']]", # 상세페이지 링크 (listBody1 내부)
                ]
                
                for xp in xpath_candidates:
                    try:
                        rows = find_in_frames(xp)
                        if rows:
                            real = []
                            for r in rows:
                                try:
                                    txt = r.get_attribute("innerText") or ""
                                    # 공고번호 패턴이 있는지 확인 (유연하게: 8자리 이상 숫자 또는 영숫자-영숫자 조합)
                                    if re.search(r'[A-Z0-9]{4,}-[A-Z0-9]+', txt) or re.search(r'\d{8,}', txt):
                                        real.append(r)
                                except: continue
                            if real:
                                bid_rows = real
                                break
                    except: continue
                
                if bid_rows:
                    print(f"   ✅ 데이터 행 {len(bid_rows)}개 감지 완료")
                    break
                
                # 아직 데이터가 없으면 잠시 대기
                if wait_step % 2 == 0:
                    # '결과가 없습니다' 메시지 확인 (조기 종료)
                    try:
                        if "결과가 없습니다" in self.driver.page_source:
                            print("   ℹ️ 검색 결과가 없습니다.")
                            break
                    except: pass
                    print(f"   ...데이터 기다리는 중 ({wait_step+1}s)")
                time.sleep(1)

            if not bid_rows:
                # 3초 더 대기 후 재시도
                print("   ⚠️ 즉시 탐색 실패, 3초 대기 후 재시도...")
                time.sleep(3)
                for xp in xpath_candidates:
                    rows = find_in_frames(xp)
                    real = []
                    for r in rows:
                        try:
                            txt = r.text
                            if re.search(r'[A-Z0-9]{4,}-[A-Z0-9]+', txt) or re.search(r'\d{8,}', txt):
                                if r.find_elements(By.TAG_NAME, "a"):
                                    real.append(r)
                        except: continue
                    if real:
                        bid_rows = real
                        print(f"   [디버그] 재시도 성공: {len(bid_rows)}개 행")
                        break

            print(f"   [디버그] 최종 탐색 행 수: {len(bid_rows)}")

            if not bid_rows:
                print("   ⚠️ 공고 데이터를 찾지 못했습니다.")

            # 매칭 시도
            print(f"📌 매칭 시도...")
            # 국방부 공고 여부 확인
            is_mnd = "국방부" in str(raw_num)
            target_row = None
            cancel_row = None
            matched_rows = []  # 매칭된 모든 행 저장 (최신 버전 선택용)
            
            for idx, row in enumerate(bid_rows):
                try:
                    row_text = row.text.replace("\n", " ").strip()
                    row_html = row.get_attribute("innerHTML")
                except:
                    continue
                
                # [수정] 국방부 공고는 첫 번째 행을 우선 선택 (번호 불일치 대응)
                if is_mnd and idx == 0:
                    print(f"      [국방부] 첫 번째 항목 자동 매칭: {row_text[:60]}...")
                    target_row = row
                    break

                # 공고번호 또는 공고명 매칭 (공백/특수문자 제거 후 비교)
                short_name = name[:10]
                name_only_ko = re.sub(r'[^가-힣0-9]', '', name)
                row_text_ko = re.sub(r'[^가-힣0-9]', '', row_text)

                is_match = False
                match_type = None
                
                # 공고번호 매칭: 정확한 번호 비교 (일련번호까지 포함)
                if clean_num and clean_num in row_text:
                    print(f"      [매칭] 공고번호 일치: {clean_num}")
                    is_match = True
                    match_type = "bid_no"
                    matched_rows.append((row, row_text, match_type))
                elif short_name and short_name in row_text:
                    print(f"      [매칭] 공고명 부분 일치: {short_name}")
                    is_match = True
                    match_type = "name"
                elif name_only_ko[:8] and name_only_ko[:8] in row_text_ko:
                    print(f"      [매칭] 공고명(한글) 유사 일치: {name_only_ko[:8]}")
                    is_match = True
                    match_type = "name_ko"
                else:
                    # 아주 짧은 매칭 시도 (실패 가능성 높음)
                    if name_only_ko[:5] and name_only_ko[:5] in row_text_ko:
                         print(f"      [매칭] 공고명(한글) 최소 일치 시도: {name_only_ko[:5]}")
                         is_match = True
                         match_type = "name_ko_min"

                if is_match:
                    print(f"   🔎 매칭 후보 발견: {row_text[:60]}...")
                    
                    is_canceled = (
                        "(취소)" in row_text
                        or 'alt="취"' in row_html
                        or 'alt="결"' in row_html
                        or "공고취소" in row_text.replace(" ", "")
                    )
                    if is_canceled:
                        print("      → 취소/마감된 공고입니다.")
                        if cancel_row is None: cancel_row = row
                        continue
                    
                    # 공고번호 매칭인 경우: 모든 매칭 행을 저장
                    if match_type == "bid_no":
                        # 나중에 일련번호가 가장 큰 것을 선택할 것임
                        target_row = None  # 계속 찾기
                    else:
                        # 공고명 매칭은 첫 번째만 사용
                        target_row = row
                        break
            
            # 공고번호 매칭된 행이 여러 개인 경우: 일련번호가 가장 큰 것 선택
            if not target_row and matched_rows:
                print(f"   [재공고 대응] 매칭된 공고 {len(matched_rows)}개 중 최신 버전 선택...")
                # 일련번호 추출 및 비교
                def extract_serial(row_text):
                    m = re.search(r'([A-Z0-9]+-(\d{3}))', row_text)
                    if m:
                        return int(m.group(2))
                    return -1
                
                matched_rows.sort(key=lambda x: extract_serial(x[1]), reverse=True)
                target_row = matched_rows[0][0]
                print(f"      → 선택됨: {matched_rows[0][1][:60]}...")

            final_row = target_row or cancel_row

            if final_row:
                try:
                    link = final_row.find_element(By.TAG_NAME, "a")
                    print(f"✅ 일치 항목 발견 → 클릭")
                    self.driver.execute_script("arguments[0].click();", link)
                    
                    # 새 창이 열릴 때까지 최대 10초 대기
                    for _ in range(20):
                        if len(self.driver.window_handles) > 1:
                            break
                        time.sleep(0.5)
                    
                    if len(self.driver.window_handles) > 1:
                        self.driver.switch_to.window(self.driver.window_handles[-1])
                    
                    # 상세 페이지 완전 로드 대기 (readyState = complete)
                    print("   ⏳ 상세 페이지 로드 대기 중...")
                    for _ in range(20):
                        try:
                            state = self.driver.execute_script("return document.readyState")
                            if state == "complete":
                                break
                        except: pass
                        time.sleep(0.5)
                    
                    # 핵심 데이터 영역이 DOM에 나타날 때까지 추가 대기
                    for _ in range(10):
                        try:
                            elems = self.driver.find_elements(By.XPATH, "//th[contains(text(), '기초금액') or contains(text(), '발주처 공고번호')]")
                            if elems:
                                break
                        except: pass
                        time.sleep(0.5)
                    
                    print("   ✅ 상세 페이지 로드 완료")
                    return True
                except Exception as e:
                    print(f"⚠️ 클릭 실패: {e}")

            # 실패 시 디버깅용 파일 저장
            print(f"🔍 입찰공고를 찾지 못했습니다.")
            try:
                os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
                with open(f"{KbidConfig.DEBUG_HISTORY_FOLDER}/search_result_debug.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                print(f"💾 search_result_debug.html 저장 완료")
            except: pass
            print(f"❌ '{name}'에 해당하는 입찰공고를 찾을 수 없습니다.")
            
        except Exception as e:
            print(f"❌ 검색 과정 중 오류 발생: {e}")
            traceback.print_exc()
        finally:
            # 상세 페이지 파싱을 위해 암묵적 대기 2초로 복구
            self.driver.implicitly_wait(2)
        return False


class KbidParser:
    """HTML 파싱 및 데이터 추출"""
    def __init__(self, driver):
        self.driver = driver

    def get_val(self, label):
        """th/td 라벨 기준 데이터 추출 (유연한 매칭 및 숫자 보정 포함)"""
        if not label: return ""
        
        # 1. 시도할 XPath 목록 (normalize-space()로 공백 무시, th/td 모두 탐색)
        # label이 "입찰(개찰) 일시" 처럼 복잡할 경우를 대비해 contains 매칭 사용
        xpaths = [
            f"(//th|//td)[contains(normalize-space(.), '{label}')][last()]/following-sibling::td[1]",
            f"(//th|//td)[contains(., '{label}')][last()]/following-sibling::td[1]",
        ]
        
        # "입찰(개찰) 일시" 특수 처리
        if "(" in label and ")" in label:
            core = label.split(")")[-1].strip() # " 일시"
            if core:
                xpaths.append(f"(//th|//td)[contains(normalize-space(.), '{core}')][last()]/following-sibling::td[1]")

        for xpath in xpaths:
            try:
                # 0.5초 짧은 대기 (전체 implicitly_wait 2초 중 일부 사용)
                element = self.driver.find_element(By.XPATH, xpath)
                text = (element.get_attribute("textContent") or element.text).replace("복사", "").strip()
                
                # 만약 가져온 텍스트가 비어있거나 '로딩중' 이면 다음 XPath 시도
                if not text or "로딩" in text:
                    continue
                    
                # 숫자가 없는 경우 input value 확인 (금액 필드 등)
                if any(k in label for k in ["금액", "가", "율"]) and not re.search(r"\d", text):
                    try:
                        val = element.find_element(By.XPATH, ".//input[@value]").get_attribute("value")
                        if val: text = val + (" 원" if "원" in text else "")
                    except: pass
                
                return re.sub(r"\s+", " ", text)
            except:
                continue
        return ""

    def get_agency(self):
        """공고기관 정보 추출 (img 태그 제거)"""
        try:
            element = self.driver.find_element(By.XPATH, KbidConfig.SELECTORS["bid_agency"])
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

    def get_val_base_price(self):
        """기초금액 전용 추출 (다중 패턴 재시도 포함)"""
        # 패턴 1: 기본 th/td 방식
        result = self.get_val("기초금액")
        if result and re.search(r"\d", result):
            return result

        print("   [기초금액] 기본 패턴 실패 → 대체 패턴 시도...")

        # 패턴 2: th 텍스트가 정확히 '기초금액'인 경우 (공백/BR 포함)
        alt_xpaths = [
            "//th[normalize-space(.)='기초금액']/following-sibling::td[1]",
            "//th[contains(text(),'기초금액')]/following-sibling::td[1]",
            "//td[normalize-space(.)='기초금액']/following-sibling::td[1]",
            "//td[contains(text(),'기초금액')]/following-sibling::td[1]",
            "//*[contains(@class,'base') and contains(text(),'기초')]/following-sibling::*[1]",
            "//label[contains(text(),'기초금액')]/following-sibling::*[1]",
        ]
        for xpath in alt_xpaths:
            try:
                elems = self.driver.find_elements(By.XPATH, xpath)
                for elem in elems:
                    text = (elem.get_attribute("textContent") or elem.text).replace("복사", "").strip()
                    text = re.sub(r"\s+", " ", text)
                    if text and re.search(r"\d", text):
                        print(f"   [기초금액] 대체 패턴 성공: {xpath} → '{text}'")
                        return text
            except: continue

        # 패턴 3: 페이지 전체 텍스트에서 정규식으로 추출
        try:
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            match = re.search(r"기초금액\s*[:\s]*([0-9,]+)\s*원?", body_text)
            if match:
                text = match.group(1)
                print(f"   [기초금액] 전체 텍스트 추출 성공: '{text}'")
                return text
        except: pass

        print("   [기초금액] 모든 패턴 실패 → 빈값 반환")
        return ""

    def get_val_a_price(self):
        """A값 추출 (합산(A) 행에서)"""
        # 패턴 1: 기본 th/td 방식으로 "합산(A)" 찾기
        alt_xpaths = [
            "//th[contains(text(), '합산(A)')]/following-sibling::td[1]",
            "//th[normalize-space(.)='합산(A)']/following-sibling::td[1]",
            "//td[contains(text(), '합산(A)')]/following-sibling::td[1]",
            "(//th|//td)[contains(text(), '합산(A)')][last()]/following-sibling::td[1]",
        ]
        
        for xpath in alt_xpaths:
            try:
                elems = self.driver.find_elements(By.XPATH, xpath)
                for elem in elems:
                    text = elem.text.strip()
                    if text and re.search(r"\d", text):
                        return text
            except:
                continue
        
        # 패턴 2: 페이지 전체 텍스트에서 정규식으로 추출
        try:
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            match = re.search(r"합산\(A\)\s*[:\s]*([0-9,]+)\s*원?", body_text)
            if match:
                return match.group(1)
        except:
            pass
        
        return ""

    def _parse_datetime(self, text):
        if not text:
            return None

        text = re.sub(r'["\u00A0]', ' ', text)
        text = text.replace('년', '-').replace('월', '-').replace('일', ' ').replace('시', ':').replace('분', ' ')
        text = re.sub(r'\s+', ' ', text).strip()

        patterns = [
            r'(\d{4}[./\-]\d{1,2}[./\-]\d{1,2}\s+\d{1,2}[:：]\d{1,2})',
            r'(\d{4}[./\-]\d{1,2}[./\-]\d{1,2})',
            r'(\d{2}[./\-]\d{1,2}[./\-]\d{1,2}\s+\d{1,2}[:：]\d{1,2})',
            r'(\d{2}[./\-]\d{1,2}[./\-]\d{1,2})'
        ]

        for pat in patterns:
            match = re.search(pat, text)
            if match:
                date_text = re.sub(r'[^0-9:\- ]', ' ', match.group()).strip()
                parts = re.findall(r'\d+', date_text)
                if len(parts) >= 5:
                    year = int(parts[0])
                    if year < 100:
                        year += 2000
                    month, day, hour, minute = map(int, parts[1:5])
                    return datetime(year, month, day, hour, minute)
                if len(parts) >= 3:
                    year = int(parts[0])
                    if year < 100:
                        year += 2000
                    month, day = map(int, parts[1:3])
                    return datetime(year, month, day, 0, 0)
        return None

    def _get_bid_status(self, deadline_text, open_time_text):
        now = datetime.now()
        deadline = self._parse_datetime(deadline_text)
        open_time = self._parse_datetime(open_time_text)

        if deadline and now < deadline:
            return "투찰대기"
        if open_time and now < open_time:
            return "투찰완료"
        if open_time and now >= open_time:
            return "결과확인"
        return "투찰대기"

    def _format_amount(self, text):
        """금액 텍스트에서 숫자만 추출하여 1,000 단위 표시 (원 제거)"""
        if not text: return ""
        
        # 1. 숫자만 추출 (콤마, 한글, 특수문자 제거)
        digits = re.sub(r"[^0-9]", "", str(text))
        
        if not digits: return ""
        
        try:
            # 2. 정수로 변환 후 콤마 추가
            val = int(digits)
            return format(val, ',')
        except:
            return digits

    def _format_rate(self, text):
        """사정률 등에서 % 제거"""
        if not text: return ""
        return str(text).replace("%", "").strip()

    def _parse_float(self, text):
        """텍스트에서 숫자(실수 포함)만 추출하여 float로 반환"""
        if not text: return None
        match = re.search(r"(\d+(?:\.\d+)?)", text)
        try:
            return float(match.group(1)) if match else None
        except: return None

    def _find_summary_value(self, label):
        """페이지 내에서 라벨 옆의 값을 찾습니다 (테이블 구조 우선)"""
        try:
            xpath = f"//th[contains(., '{label}')]/following-sibling::td[1] | //td[contains(., '{label}')]/following-sibling::td[1]"
            elements = self.driver.find_elements(By.XPATH, xpath)
            for elem in elements:
                text = elem.text.strip()
                if text: return text
            
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            pattern = rf"{re.escape(label)}\s*[:：]?\s*([0-9]+[.,]?[0-9]*\s*%?)"
            match = re.search(pattern, body_text)
            return match.group(1).strip() if match else ""
        except: return ""

    def _parse_result_table(self):
        """낙찰순위 테이블에서 행 및 헤더를 추출합니다."""
        try:
            # 1. 다양한 XPath 패턴으로 테이블 찾기
            table = None
            xpath_patterns = [
                "//div[contains(., '낙찰순위')]//table[1]",
                "//table[.//th[contains(., '순위')] and .//th[contains(., '상호명')]]",
                "//table[.//th[contains(., '순위')] and .//th[contains(., '투찰률')]]",
                "//table[@class='table' or @class='tbl']//tr[.//th]",
                "//div[@id='listBody1']//table",
                "//div[contains(@class, 'list')]//table"
            ]
            
            for xpath in xpath_patterns:
                try:
                    table = self.driver.find_element(By.XPATH, xpath)
                    if table:
                        break
                except:
                    continue
            
            if not table:
                print("⚠️ 낙찰순위 테이블을 찾을 수 없습니다.")
                return None, []
            
            # 2. 헤더 추출
            headers = []
            try:
                header_cells = table.find_elements(By.XPATH, ".//th")
                headers = [h.text.strip() for h in header_cells]
            except:
                # th가 없으면 첫 번째 tr을 헤더로 간주
                try:
                    first_row = table.find_element(By.XPATH, ".//tr[1]")
                    cells = first_row.find_elements(By.XPATH, ".//td | .//th")
                    headers = [c.text.strip() for c in cells]
                except:
                    pass
            
            if not headers:
                print("⚠️ 테이블 헤더를 추출할 수 없습니다.")
                return None, []
            
            print(f"📋 테이블 헤더: {headers}")
            
            # 3. 행 데이터 추출
            rows = []
            try:
                row_elements = table.find_elements(By.XPATH, ".//tr[td]")
                for row_elem in row_elements:
                    cells = [c.text.strip() for c in row_elem.find_elements(By.XPATH, ".//td")]
                    if not cells:
                        continue
                    
                    # 헤더 개수와 맞추기
                    if len(cells) < len(headers):
                        cells = cells + [""] * (len(headers) - len(cells))
                    elif len(cells) > len(headers):
                        cells = cells[:len(headers)]
                    
                    row_dict = {headers[i]: cells[i] for i in range(len(headers))}
                    rows.append(row_dict)
                    print(f"   행 {len(rows)}: {row_dict}")
            except Exception as e:
                print(f"⚠️ 테이블 행 추출 중 오류: {e}")
            
            print(f"✅ {len(rows)}개 행 추출 완료")
            return headers, rows
        
        except Exception as e:
            print(f"❌ 낙찰순위 테이블 파싱 오류: {e}")
            return None, []
    

    def _get_company_result(self, company_name, headers, rows):
        for row in rows:
            for key, value in row.items():
                if company_name in value:
                    return row
        return None

    def _parse_result_details(self):
        """결과 페이지에서 두 투찰 회사 및 요약값을 추출합니다."""
        headers, rows = self._parse_result_table()
        result_data = {
            "참여 업체수": "", "사정률": "", "1등 상호명": "",
            "1등 업체 입찰금액": "", "1등 업체 사정률": "",
            "AIR 채호원 입찰금액": "", "AIR 채호원 사정률": "", "AIR 채호원 순위": "", 
            "에어채호원 입찰금액": "", "에어채호원 사정률": "", "에어채호원 순위": "",
        }

        result_data["사정률"] = self._format_rate(self._find_summary_value("사정률"))
        if not result_data["사정률"]:
             val = self._find_summary_value("낙찰율")
             if val: result_data["사정률"] = self._format_rate(val)

        if not rows or not headers: return result_data
        result_data["참여 업체수"] = str(len(rows))

        # 콜럼 인덱스 분리 (카테고리별)
        rank_idx = company_idx = bid_amount_idx = corp_rate_idx = -1
        for i, h in enumerate(headers):
            if "순위" in h:       rank_idx = i
            if "상호" in h:       company_idx = i
            if "입찰금액" in h:   bid_amount_idx = i
            if "사정률" in h and ("업체" in h or "사정률" in h): corp_rate_idx = i

        # 두 회사 이름 (공백 제거 후 비교)
        company_air  = "AIR채호원"       # AIR 채호원
        company_corp = "주식회사에어채호원"   # 주식회사 에어채호원

        for idx, row in enumerate(rows):
            row_vals = list(row.values())
            if idx == 0:
                if corp_rate_idx != -1:
                    result_data["1등 업체 사정률"] = self._format_rate(row_vals[corp_rate_idx])
                if bid_amount_idx != -1:
                    result_data["1등 업체 입찰금액"] = self._format_amount(row_vals[bid_amount_idx])
                if company_idx != -1:
                    result_data["1등 상호명"] = row_vals[company_idx]
            if company_idx == -1:
                continue
            name_clean = row_vals[company_idx].replace(" ", "")
            if company_air in name_clean:
                if rank_idx != -1:       result_data["AIR 채호원 순위"] = row_vals[rank_idx]
                if bid_amount_idx != -1: result_data["AIR 채호원 입찰금액"] = self._format_amount(row_vals[bid_amount_idx])
                if corp_rate_idx != -1:  result_data["AIR 채호원 사정률"] = self._format_rate(row_vals[corp_rate_idx])
            if company_corp in name_clean:
                if rank_idx != -1:       result_data["에어채호원 순위"] = row_vals[rank_idx]
                if bid_amount_idx != -1: result_data["에어채호원 입찰금액"] = self._format_amount(row_vals[bid_amount_idx])
                if corp_rate_idx != -1:  result_data["에어채호원 사정률"] = self._format_rate(row_vals[corp_rate_idx])
        return result_data

    def parse_all(self):
        """상세 공고 정보 추출"""
        bid_close_text = re.sub(r"투찰(?:하기)?", "", self.get_val("투찰마감일시")).strip()
        open_time_text = self.get_val("입찰(개찰) 일시")
        
        # [수정] 개찰일시가 없는 경우 투찰마감일시 기준 +1시간 자동 설정
        if not open_time_text or not re.search(r"\d", open_time_text):
            deadline_dt = self._parse_datetime(bid_close_text)
            if deadline_dt:
                fallback_dt = deadline_dt + timedelta(hours=1)
                open_time_text = fallback_dt.strftime("%Y-%m-%d %H:%M")
                print(f"   ℹ️ 개찰일시 누락 → 투찰마감(+1h)으로 자동 설정: {open_time_text}")
        
        # 공고번호 추출: 일련번호까지 포함해서 추출 (조달청은 -002 등 재공고 구분)
        bid_no = ""
        
        # 1단계: "정정/취소 공고" 섹션에서 현재 공고(빨간색) 찾기 (재공고 대응)
        try:
            # 정정/취소 공고 섹션에서 font color=red인 항목이 현재 공고
            current_bid_elem = self.driver.find_element(
                By.XPATH,
                "//h4[contains(text(), '정정/취소 공고')]/following-sibling::div[1]//font[@color='red']"
            )
            current_bid_text = current_bid_elem.text.strip() if current_bid_elem else ""
            if current_bid_text:
                bid_no = current_bid_text
                print(f"   [디버그] '정정/취소 공고' 섹션에서 현재 공고(빨간색) 추출: {bid_no}")
        except:
            pass
        
        # 2단계: 실패 시 페이지 텍스트에서 "조달청 시설 R26BK01526419-002 호" 형태로 추출
        if not bid_no:
            try:
                page_text = self.driver.page_source
                # "조달청 시설 R26BK01526419-002 호" 패턴으로 추출
                m = re.search(r'([A-Z0-9]+-\d{3})\s+호', page_text)
                if m:
                    bid_no = m.group(1).strip()
                    print(f"   [디버그] 페이지 텍스트에서 공고번호 추출: {bid_no}")
            except:
                pass
        
        # 3단계: 실패 시 copythis 버튼에서 추출 (일련번호 없음, fallback)
        if not bid_no:
            try:
                copy_btn = self.driver.find_element(
                    By.XPATH,
                    "(//th[contains(text(), '발주처 공고번호')])[last()]/following-sibling::td[1]//input[contains(@onclick, 'copythis')]"
                )
                onclick_val = copy_btn.get_attribute("onclick") or ""
                m = re.search(r"copythis\('([^']+)'\)", onclick_val)
                if m:
                    bid_no = m.group(1).strip()
                    print(f"   [디버그] copythis 버튼에서 공고번호 추출(일련번호 없음): {bid_no}")
            except:
                pass
        
        # 4단계: 실패 시 텍스트 필드에서 추출 (최후의 수단)
        if not bid_no:
            bid_no = self.get_val("발주처 공고번호")
        if not bid_no:
            bid_no = self.get_val("공고번호")

        # 디버깅을 위해 상세 페이지 HTML 저장
        try:
            os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
            safe_bid_no = bid_no.replace("/", "_").replace("\\", "_") if bid_no else "unknown"
            html_filename = f"{KbidConfig.DEBUG_HISTORY_FOLDER}/detail_{safe_bid_no}.html"
            with open(html_filename, "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
            print(f"   [디버그] 상세 페이지 HTML 저장: {html_filename}")
        except Exception as e:
            print(f"   ⚠️ HTML 저장 실패: {str(e)[:50]}")

        
        data = {
            "투찰상태": self._get_bid_status(bid_close_text, open_time_text),
            "공고번호": bid_no,
            "공고기관": self.get_agency(),
            "공고명": "",
            "지역제한": self.get_val("지역제한"),
            "업종": self.get_val("업종"),
            "입찰개시일": self.get_val("입찰개시일"),
            "투찰마감일시": bid_close_text,
            "개찰일시": open_time_text,
            "예상투찰가1": "",
            "예상투찰가2": "",
            "예상투찰가3": "",
            "기초금액": self._format_amount(self.get_val_base_price()),
            "A값": self._format_amount(self.get_val_a_price()),
            "예가변동폭": self.get_val("예가변동폭"),
            "투찰하한율": self.get_val("투찰하한율"),
            "계약방법": self.get_val("계약방법")
        }
        
        # [수정] 예상투찰가 자동 계산
        try:
            # 콤마 제거 후 float 변환
            clean_base = data["기초금액"].replace(",", "") if data["기초금액"] else ""
            base_val = float(clean_base) if clean_base else None
            
            clean_a = data["A값"].replace(",", "") if data["A값"] else ""
            a_val = float(clean_a) if clean_a else None
            
            limit_val = self._parse_float(data["투찰하한율"])
            rate_text = self.get_val("평균사정률") or self.get_val("사정률")
            rate_val = self._parse_float(rate_text)
            
            # A값이 있을 때: (기초금액*사정률-A)*투찰하한율+A
            if a_val and a_val > 0:
                if base_val and rate_val and limit_val:
                    # 예상투찰가1과 예상투찰가2 모두 A값 공식 적용
                    est = (base_val * (rate_val / 100.0) - a_val) * (limit_val / 100.0) + a_val
                    data["예상투찰가1"] = format(int(est), ',')
                    data["예상투찰가2"] = format(int(est), ',')
                    print(f"   💰 예상투찰가 자동계산 완료 (A값 포함): {data['예상투찰가1']} 공식: (기초금액*사정률-A)*투찰하한율+A")
                else:
                    data["예상투찰가1"] = "-"
                    data["예상투찰가2"] = "-"
            
            # A값이 없을 때: 기존 방식
            else:
                # 예상투찰가1: 기초금액 * 투찰하한율
                if base_val and limit_val:
                    est1 = base_val * (limit_val / 100.0)
                    data["예상투찰가1"] = format(int(est1), ',')
                    print(f"   💰 예상투찰가1 자동계산 완료: {data['예상투찰가1']} (하한율: {limit_val}%)")
                else:
                    data["예상투찰가1"] = "-"

                # 예상투찰가2: 기초금액 * 평균 사정률 * 투찰하한율
                if base_val and rate_val and limit_val:
                    # 공식: 기초금액 * (사정률/100) * (하한율/100)
                    est2 = base_val * (rate_val / 100.0) * (limit_val / 100.0)
                    data["예상투찰가2"] = format(int(est2), ',')
                    print(f"   💰 예상투찰가2 자동계산 완료: {data['예상투찰가2']} (사정률: {rate_val}%)")
                else:
                    data["예상투찰가2"] = "-"
        except Exception as e:
            data["예상투찰가1"] = "-"
            data["예상투찰가2"] = "-"
            print(f"   ⚠️ 예상투찰가 계산 실패: {e}")
        
        try:
            # JS를 사용해 DOM 조작으로 불필요한 태그(img, 아이콘 span 등) 제거 후 텍스트 추출
            script = """
                var elem = document.querySelector(arguments[0]);
                if (!elem) return '';
                var clone = elem.cloneNode(true);
                // 1. img 태그(접두사 아이콘 등) 제거
                var imgs = clone.querySelectorAll('img');
                imgs.forEach(img => img.remove());
                
                // 2. 공고 종류 표시 span 아이콘 제거 (예: <span class="ico-bid kind5">전</span>) 
                // span 태그만 제거하고 텍스트 노드는 보존하도록 함
                var iconSpans = clone.querySelectorAll('span[class*="ico-bid"], span[class*="kind"]');
                iconSpans.forEach(span => {
                    while (span.firstChild) {
                        span.parentNode.insertBefore(span.firstChild, span);
                    }
                    span.parentNode.removeChild(span);
                });
                
                // 3. 투찰마감 등 특정 텍스트를 가진 독립된 태그 제거
                // (주의: '전' 단독 제거는 제거함 - 공고명의 첫 글자가 '전기', '전자' 등인 경우 보호)
                var els = clone.querySelectorAll('*');
                var removeTexts = ['투찰마감', '결과확인', '투찰대기', '입찰마감', '투찰완료', '긴', '결', '수', '취', '견', '재', '전자', '긴급', '취소', '과거 공고', '지난공고', '지나온공고'];
                els.forEach(el => {
                    if (removeTexts.includes(el.textContent.trim())) {
                        el.remove();
                    }
                });
                return clone.textContent;
            """
            raw_title = self.driver.execute_script(script, KbidConfig.SELECTORS["detail_title"])
            
            if "(취소)" in raw_title:
                data["투찰상태"] = "공고 취소"
                
            clean_title = raw_title
            # 텍스트로 섞여 있을 수 있는 상태값 제거
            for word in ['투찰마감', '결과확인', '투찰대기', '입찰마감', '투찰완료']:
                clean_title = clean_title.replace(word, "")
                
            # 괄호로 둘러싸인 접두사/접미사 제거 ('전' 제거 안 함 - 공고명이 '전기' 등인 경우 보호)
            clean_title = re.sub(r'[\[\(<](?:긴|취|긴급|전자|취소|지난공고|지나온공고|공고|재공고|재공고입찰|견적)[\]\)>]', '', clean_title)
            clean_title = clean_title.replace("(취소)", "").replace("(과거 공고)", "")
            clean_title = clean_title.replace("(재공고)", "").replace("[재공고]", "").replace("<재공고입찰>", "").replace("(견적)", "")
            
            # 띄어쓰기가 있는 접두사 제거 (예: "긴 용인..." -> "용인...") - '전' 제거 안 함
            clean_title = re.sub(r'^(?:긴|취|급|수|결|견|재)\s+', '', clean_title).replace("복사", "")
            
            data["공고명"] = re.sub(r'\s+', ' ', clean_title).strip()
        except: pass
        return data

class KbidCrawler:
    """프로세스 통합 컨트롤러"""
    def __init__(self):
        try:
            self.gs = GoogleSheetsManager()
            self.browser = None
            self.parser = None
        except Exception as e:
            print(f"❗ 초기화 오류: {e}")
            exit()

    def _remove_duplicate_bids(self):
        """공고명 + 공고번호 기준으로 중복 및 재공고 처리 (최신 데이터만 유지)"""
        try:
            print("\n🔍 중복/재공고 처리 중...")
            ws = self.gs.sheet.worksheet("투찰준비")
            headers = [h.replace("*", "").strip() for h in ws.row_values(1)]
            bid_name_col_idx = headers.index("공고명") if "공고명" in headers else 1
            bid_no_col_idx = headers.index("공고번호") if "공고번호" in headers else 2
            
            all_data = ws.get_all_values()
            header_row = all_data[0]
            data_rows = all_data[1:]
            
            if not data_rows:
                print("✅ 데이터가 없어 중복 제거 완료")
                return
            
            # 공고명별로 그룹화하고, 각 공고명 내에서 공고번호 숫자가 가장 큰 행만 유지
            bid_groups = {}  # {공고명: [(인덱스, 공고번호, 숫자값)]}
            
            def extract_bid_number(bid_no_str):
                """공고번호에서 숫자만 추출"""
                if not bid_no_str:
                    return -1
                # 숫자만 추출
                digits = re.sub(r'[^0-9]', '', bid_no_str)
                try:
                    return int(digits) if digits else -1
                except:
                    return -1
            
            for idx, row in enumerate(data_rows):
                bid_name = row[bid_name_col_idx].strip() if len(row) > bid_name_col_idx else ""
                bid_no = row[bid_no_col_idx].strip() if len(row) > bid_no_col_idx else ""
                
                if bid_name:
                    # 공고명을 정규화 (공백 정리)
                    normalized_name = re.sub(r'\s+', ' ', bid_name).strip()
                    
                    if normalized_name not in bid_groups:
                        bid_groups[normalized_name] = []
                    
                    bid_num_value = extract_bid_number(bid_no)
                    bid_groups[normalized_name].append((idx, bid_no, bid_num_value))
            
            # 유지할 행 결정 (각 공고명별로 숫자값이 가장 큰 것만)
            rows_to_keep_indices = set()
            rows_to_delete = []
            
            for bid_name, entries in bid_groups.items():
                if len(entries) == 1:
                    # 공고명이 유일하면 유지
                    rows_to_keep_indices.add(entries[0][0])
                else:
                    # 여러 개라면 숫자값이 가장 큰 것만 유지
                    entries_sorted = sorted(entries, key=lambda x: x[2], reverse=True)
                    latest_idx = entries_sorted[0][0]
                    latest_bid_no = entries_sorted[0][1]
                    
                    rows_to_keep_indices.add(latest_idx)
                    
                    # 나머지는 삭제 대상
                    for idx, bid_no, bid_num in entries_sorted[1:]:
                        rows_to_delete.append((idx, bid_no, latest_bid_no, bid_name))
            
            # 최종 행 구성 (유지할 행만)
            rows_to_keep = [data_rows[idx] for idx in sorted(rows_to_keep_indices)]
            
            # 삭제 로그 출력
            if rows_to_delete:
                print(f"   🗑️ 다음 공고들을 삭제합니다 (최신 공고로 대체됨):")
                for old_idx, old_bid_no, new_bid_no, name in rows_to_delete:
                    print(f"      - {old_bid_no} → {new_bid_no} ({name[:40]}...)")
            
            duplicate_count = len(data_rows) - len(rows_to_keep)
            
            if duplicate_count > 0:
                # 시트 업데이트 (헤더 + 유지할 행들)
                new_data = [header_row] + rows_to_keep
                ws.clear()
                ws.update(values=new_data, range_name="A1")
                print(f"✅ {duplicate_count}건의 중복/재공고를 제거했습니다.")
            else:
                print("✅ 중복/재공고가 없습니다.")
        
        except Exception as e:
            print(f"⚠️ 중복 제거 중 오류 발생: {e}")
            traceback.print_exc()


    def run(self):
        try:
            # 1. 투찰준비 시트의 상태를 최신화
            updated_count = self.gs.update_bid_statuses()
            if updated_count is not None:
                print(f"⚙️ 투찰준비 시트 상태 확인 완료 ({updated_count}건 업데이트)")

            # 2. 시트에서 공고 목록 및 이미 처리된 목록 로드
            tasks = self.gs.get_search_terms()
            processed_bids = self.gs.get_processed_bids()
            
            if not tasks:
                print("📝 작업할 공고가 없습니다. 투찰준비 시트 상태를 먼저 확인했습니다.")
                return

            # [최적화] 기초금액 누락 공고 목록 사전 조회 (로그인 전에 처리)
            # 이것은 로그인이 필요 없으므로 여기서 처리하여 로그인 후 지연 방지
            try:
                ws_prepare = self.gs.sheet.worksheet("투찰준비")
                headers_row = [h.replace("*", "").strip() for h in ws_prepare.row_values(1)]
                base_price_col_idx = headers_row.index("기초금액") if "기초금액" in headers_row else -1
                bid_no_col_idx = headers_row.index("공고번호") if "공고번호" in headers_row else 1
                all_rows = ws_prepare.get_all_values()[1:]  # 헤더 제외
                missing_base_price_bids = set()
                for row in all_rows:
                    bid_no_val = row[bid_no_col_idx].strip() if len(row) > bid_no_col_idx else ""
                    base_price_val = row[base_price_col_idx].strip() if base_price_col_idx >= 0 and len(row) > base_price_col_idx else ""
                    if bid_no_val and not base_price_val:
                        missing_base_price_bids.add(bid_no_val)
                if missing_base_price_bids:
                    print(f"⚠️ 기초금액 누락 공고 {len(missing_base_price_bids)}건 재확인 예정")
            except Exception as e:
                print(f"⚠️ 기초금액 누락 조회 중 오류 (무시): {e}")
                missing_base_price_bids = set()

            self.browser = KbidBrowser(self.gs)
            self.parser = KbidParser(self.browser.driver)
            
            if not tasks:
                print("📝 작업할 공고가 없습니다.")
                return

            # [최적화] 로그인 수행 (사전 조회 완료 후 빠르게 진행)
            print("\n🔐 로그인 절차를 시작합니다...")
            if not self.browser.wait_for_login(): return

            # 3. 공고 순회 크롤링
            for task in tasks:
                try:
                    name = task.get("name", "").strip()
                    num = task.get("num", "").strip()
                    display_name = num or name
                    
                    # 핵심 번호만 추출하여 중복 체크 (예: R26BK01488658-000)
                    match = re.search(r'[A-Z0-9]{5,}-[A-Z0-9]+', num)
                    clean_num = match.group() if match else num
                    
                    if clean_num and clean_num in processed_bids:
                        # 기초금액이 누락된 경우에는 재크롤링 시도
                        if clean_num in missing_base_price_bids:
                            print(f"🔄 기초금액 누락으로 재크롤링: {clean_num}")
                        else:
                            print(f"⏩ 이미 시트에 있는 공고이므로 건너뜁니다: {clean_num}")
                            continue

                    print(f"\n🔍 작업 확인: {display_name}")
                    
                    # 상세 페이지 이동 (공고 정보 기반)
                    if self.browser.navigate_to_bid(task):
                        res_data = self.parser.parse_all()
                        
                        # 공고명은 입찰공고 시트의 원본 이름을 그대로 사용
                        # (웹 스크랩 시 kbid 배지 텍스트(전, 긴, A 등)가 섞이는 문제 방지)
                        if name:
                            res_data["공고명"] = name
                        
                        # 공고번호도 매칭 안정성을 위해 가급적 원본 번호를 유지
                        if num:
                            res_data["공고번호"] = num
                        
                        bid_no = res_data.get("공고번호")
                        
                        # 항상 저장을 시도 (내부 로직에서 업데이트/추가 결정)
                        if self.gs.save_result(res_data):
                            print(f"   └─ 투찰상태: {res_data.get('투찰상태')}")
                            if res_data.get('AIR 채호원 순위'):
                                print(f"   └─ AIR 채호원: {res_data.get('AIR 채호원 순위')}위 / {res_data.get('AIR 채호원 입찰금액')}")
                            if res_data.get('에어채호원 순위'):
                                print(f"   └─ 에어채호원: {res_data.get('에어채호원 순위')}위 / {res_data.get('에어채호원 입찰금액')}")  
                            if res_data.get('1등 상호명'):
                                print(f"   └─ 1등: {res_data.get('1등 상호명')} / 업체사정률 {res_data.get('1등 업체사정률')}")
                        
                        # 상세 페이지(새 창) 닫기
                        if len(self.browser.driver.window_handles) > 1:
                            self.browser.driver.close()
                            self.browser.driver.switch_to.window(self.browser.driver.window_handles[0])
                    else:
                        print(f"❌ 공고를 찾을 수 없음: {display_name}")
                        # 확인불가 상태로 시트에 기록
                        fail_data = {
                            "투찰상태": "확인불가",
                            "공고번호": clean_num or num,
                            "공고명": name,
                        }
                        self.gs.save_result(fail_data)
                        print(f"   └─ '투찰준비' 시트에 '확인불가'로 기록했습니다.")
                except Exception as e:
                    print(f"❗ '{display_name}' 처리 중 예상치 못한 에러 발생: {e}")
                    traceback.print_exc()
                
                time.sleep(0.3)
            
            # --- [추가 기능] 투찰준비 시트의 기초금액 누락 행 재탐색 ---
            print("\n🔄 투찰준비 시트의 기초금액 누락 행 재탐색을 시작합니다.")
            try:
                ws_prepare = self.gs.sheet.worksheet("투찰준비")
                headers_row = [h.replace("*", "").strip() for h in ws_prepare.row_values(1)]
                base_price_col_idx = headers_row.index("기초금액") if "기초금액" in headers_row else -1
                bid_no_col_idx = headers_row.index("공고번호") if "공고번호" in headers_row else 1
                bid_name_col_idx = headers_row.index("공고명") if "공고명" in headers_row else 2
                status_col_idx = headers_row.index("투찰상태") if "투찰상태" in headers_row else 0
                
                if base_price_col_idx >= 0:
                    all_rows = ws_prepare.get_all_values()[1:]  # 헤더 제외
                    missing_tasks = []
                    
                    # 기초금액이 누락되고 투찰대기 상태인 공고만 수집 (행 인덱스도 함께 저장)
                    for row_idx, row in enumerate(all_rows, start=2):  # 2부터 시작 (헤더는 1행)
                        bid_no_val = row[bid_no_col_idx].strip() if len(row) > bid_no_col_idx else ""
                        bid_name_val = row[bid_name_col_idx].strip() if len(row) > bid_name_col_idx else ""
                        base_price_val = row[base_price_col_idx].strip() if len(row) > base_price_col_idx else ""
                        status_val = row[status_col_idx].strip() if len(row) > status_col_idx else ""
                        
                        # 조건: 기초금액이 비어있고, 상태가 '투찰대기'일 경우
                        if (bid_no_val or bid_name_val) and not base_price_val and status_val == "투찰대기":
                            missing_tasks.append({
                                "num": bid_no_val, 
                                "name": bid_name_val, 
                                "row_idx": row_idx  # 기존 행 인덱스 저장
                            })
                    
                    if missing_tasks:
                        print(f"📝 기초금액 누락 공고 {len(missing_tasks)}건 발견. 재탐색을 진행합니다.")
                        for task in missing_tasks:
                            name = task.get("name", "").strip()
                            num = task.get("num", "").strip()
                            row_idx = task.get("row_idx")
                            display_name = num or name
                            print(f"\n🔍 [재탐색] 작업 확인: {display_name} (행 {row_idx})")
                            
                            try:
                                if self.browser.navigate_to_bid(task):
                                    res_data = self.parser.parse_all()
                                    if name:
                                        res_data["공고명"] = name
                                    if num:
                                        res_data["공고번호"] = num
                                    
                                    # save_result 메소드가 기존 행을 찾아 업데이트함
                                    if self.gs.save_result(res_data):
                                        print(f"   └─ 재탐색 기초금액 업데이트 완료: {res_data.get('기초금액')}")
                                    
                                    if len(self.browser.driver.window_handles) > 1:
                                        self.browser.driver.close()
                                        self.browser.driver.switch_to.window(self.browser.driver.window_handles[0])
                                else:
                                    print(f"❌ [재탐색] 공고를 찾을 수 없음: {display_name}")
                            except Exception as e:
                                print(f"❗ [재탐색] '{display_name}' 처리 중 에러 발생: {e}")
                                traceback.print_exc()
                            time.sleep(0.3)
                        
                        # 재탐색 완료 후 중복 제거
                        self._remove_duplicate_bids()
                    else:
                        print("✅ 기초금액이 누락된 공고가 없습니다.")
            except Exception as e:
                print(f"⚠️ 기초금액 누락 행 재탐색 중 오류 발생: {e}")
                traceback.print_exc()
            # --------------------------------------------------------
            
        except Exception as e:
            print(f"🛑 프로그램 실행 중 치명적 오류 발생")
            traceback.print_exc()
        finally:
            if self.browser is not None:
                try:
                    self.browser.driver.quit()
                except Exception:
                    pass
            print("\n🏁 모든 작업을 마쳤습니다.")

if __name__ == "__main__":
    os.makedirs(KbidConfig.DEBUG_HISTORY_FOLDER, exist_ok=True)
    KbidCrawler().run()
