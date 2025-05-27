from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
import re
from .config import Config

class GroupwareAutomation:
    """그룹웨어 자동화 클래스 - 체크박스 문제 해결"""
    
    def __init__(self):
        self.config = Config()
        self.driver = None
        self.wait = None
        self.login_url = "https://gw.meta-m.co.kr/gw/userMain.do"

    def setup_driver(self):
        """WebDriver 설정"""
        try:
            options = Options()
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument('--disable-web-security')
            options.add_argument('--allow-running-insecure-content')
            
            self.driver = webdriver.Chrome(options=options)
            self.wait = WebDriverWait(self.driver, 30)
            print("✅ 브라우저 시작 완료")
            return True
        except Exception as e:
            raise Exception(f"브라우저 실행 실패: {e}")

    def login_to_groupware(self, user_id, password):
        """그룹웨어 로그인"""
        try:
            print("🔐 로그인 시작...")
            self.driver.get(self.login_url)
            time.sleep(3)
            
            # 아이디 입력
            id_input = self.wait.until(EC.presence_of_element_located((By.ID, "userId")))
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(1)
            
            # 비밀번호 입력
            pw_input = self.driver.find_element(By.ID, "userPw")
            pw_input.clear()
            pw_input.send_keys(password)
            time.sleep(1)
            
            # 로그인 버튼 클릭
            pw_input.send_keys(Keys.ENTER)
            time.sleep(5)
            
            # 로그인 성공 확인
            current_url = self.driver.current_url
            if "userMain.do" not in current_url:
                raise Exception("로그인에 실패했습니다. 아이디와 비밀번호를 확인해주세요.")
            
            print("✅ 로그인 완료")
            return True
            
        except Exception as e:
            raise Exception(f"로그인 실패: {e}")

    def navigate_to_expense_page(self):
        """지출결의서 페이지로 이동"""
        try:
            print("🌐 지출결의서 페이지로 이동...")
            self.driver.get(self.config.GROUPWARE_URL)
            time.sleep(5)
            print("✅ 페이지 이동 완료")
            return True
        except Exception as e:
            raise Exception(f"페이지 이동 실패: {e}")

    def setup_card_interface(self, start_date, end_date):
        """카드 사용내역 인터페이스 설정 (1회만 실행)"""
        try:
            print("💳 카드 사용내역 설정 시작...")
            
            # 1. 카드 사용내역 버튼 클릭
            print("  1) 카드 사용내역 버튼 클릭")
            card_btn = self.wait.until(EC.element_to_be_clickable((By.ID, "btnExpendInterfaceCard")))
            card_btn.click()
            time.sleep(3)
            
            # 2. 선택 버튼 클릭
            print("  2) 선택 버튼 클릭")
            select_btn = self.wait.until(EC.element_to_be_clickable((By.ID, "btnExpendCardInfoHelpPop")))
            select_btn.click()
            time.sleep(3)
            
            # 3. 날짜 입력
            print(f"  3) 날짜 입력: {start_date} ~ {end_date}")
            self._input_dates(start_date, end_date)
            
            # 4. 검색 버튼 클릭
            print("  4) 검색 버튼 클릭")
            search_btn = self.wait.until(EC.element_to_be_clickable((By.ID, "btnExpendCardListSearch")))
            search_btn.click()
            
            # 5. 검색 결과 로딩 대기 (중요!)
            print("  5) 검색 결과 로딩 대기...")
            time.sleep(8)  # 더 긴 대기 시간
            
            # 6. 검색 결과 확인
            self._verify_search_results()
            
            print("✅ 카드 사용내역 설정 완료")
            return True
            
        except Exception as e:
            raise Exception(f"카드 사용내역 설정 실패: {e}")

    def _input_dates(self, start_date, end_date):
        """날짜 입력 (개선된 버전 - TAB 키 사용)"""
        try:
            # 날짜 형식 변환 (YYYYMMDD -> YYYY-MM-DD)
            formatted_start = f"{start_date[:4]}-{start_date[4:6]}-{start_date[6:8]}"
            formatted_end = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:8]}"
            
            print(f"    시작날짜 입력: {formatted_start}")
            
            # 시작 날짜 입력
            start_input = self.wait.until(EC.presence_of_element_located((By.ID, "txtExpendCardFromDate")))
            self._clear_and_input_with_tab(start_input, formatted_start)
            
            print(f"    종료날짜 입력: {formatted_end}")
            
            # 종료 날짜 입력
            end_input = self.driver.find_element(By.ID, "txtExpendCardToDate")
            self._clear_and_input_final(end_input, formatted_end)
            
            # 검증
            time.sleep(2)
            actual_start = start_input.get_attribute('value')
            actual_end = end_input.get_attribute('value')
            print(f"    입력 확인 - 시작: {actual_start}, 종료: {actual_end}")
            
            return True
            
        except Exception as e:
            # JavaScript 백업 방법
            print(f"    키보드 입력 실패, JavaScript로 재시도: {e}")
            return self._input_dates_with_javascript(formatted_start, formatted_end)

    def _clear_and_input_with_tab(self, element, value):
        """요소 클리어 후 값 입력 (TAB 키 사용)"""
        element.click()
        time.sleep(0.5)
        element.send_keys(Keys.CONTROL + "a")
        time.sleep(0.3)
        element.send_keys(Keys.DELETE)
        time.sleep(0.3)
        element.send_keys(value)
        time.sleep(0.5)
        element.send_keys(Keys.TAB)  # ENTER 대신 TAB 사용
        time.sleep(1)

    def _clear_and_input_final(self, element, value):
        """마지막 날짜 입력 (ENTER나 TAB 없이)"""
        element.click()
        time.sleep(0.5)
        element.send_keys(Keys.CONTROL + "a")
        time.sleep(0.3)
        element.send_keys(Keys.DELETE)
        time.sleep(0.3)
        element.send_keys(value)
        time.sleep(1)
        # 마지막이므로 키 입력 없이 포커스만 이동
        self.driver.find_element(By.TAG_NAME, "body").click()
        time.sleep(1)

    def _input_dates_with_javascript(self, formatted_start, formatted_end):
        """JavaScript를 사용한 날짜 입력"""
        try:
            js_script = f"""
            // 시작 날짜 설정
            var startInput = document.getElementById('txtExpendCardFromDate');
            startInput.value = "{formatted_start}";
            startInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
            
            // 종료 날짜 설정
            var endInput = document.getElementById('txtExpendCardToDate');
            endInput.value = "{formatted_end}";
            endInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
            
            return [startInput.value, endInput.value];
            """
            
            result = self.driver.execute_script(js_script)
            print(f"    JavaScript 입력 결과: 시작={result[0]}, 종료={result[1]}")
            time.sleep(2)
            return True
            
        except Exception as e:
            raise Exception(f"JavaScript 날짜 입력 실패: {e}")

    def _verify_search_results(self):
        """검색 결과 확인"""
        try:
            # 금액 테이블이 로딩될 때까지 대기
            print("    검색 결과 테이블 확인...")
            
            # 금액 셀들이 나타날 때까지 대기
            amount_cells = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td.td_ri span.fwb"))
            )
            
            print(f"    검색 결과: {len(amount_cells)}개 항목 발견")
            
            # 처음 몇 개 금액 출력해서 확인
            for i, cell in enumerate(amount_cells[:3]):  # 처음 3개만 출력
                print(f"    금액 {i+1}: {cell.text}")
            
            return True
            
        except Exception as e:
            print(f"    ⚠️ 검색 결과 확인 실패: {e}")
            return False

    def process_single_record(self, data_row, record_index, total_records):
        """단일 레코드 처리"""
        try:
            print(f"\n📝 레코드 {record_index}/{total_records} 처리 시작")
            print(f"   처리할 금액: {data_row.get('amount', '')}")
            
            # 1. 금액 매칭하여 체크박스 클릭
            print("   1) 금액 매칭 및 체크박스 클릭")
            success = self._find_and_click_checkbox(data_row.get('amount', ''))
            
            if not success:
                print(f"   ❌ 금액 매칭 실패: {data_row.get('amount', '')}")
                return False
            
            # 2. 폼 데이터 입력
            print("   2) 폼 데이터 입력")
            self._input_form_data(data_row)
            
            # 3. 저장
            print("   3) 저장")
            self._click_save()
            
            print(f"   ✅ 레코드 {record_index} 완료")
            return True
            
        except Exception as e:
            print(f"   ❌ 레코드 {record_index} 실패: {e}")
            return False

    def _find_and_click_checkbox(self, target_amount):
        """금액 매칭하여 체크박스 클릭 (개선된 버전)"""
        try:
            clean_target = self._clean_amount(str(target_amount))
            print(f"      찾는 금액: {clean_target}")
            
            # 새로고침 후 금액 셀들 다시 찾기
            time.sleep(2)
            amount_cells = self.driver.find_elements(By.CSS_SELECTOR, "td.td_ri span.fwb")
            
            if not amount_cells:
                print("      ❌ 금액 셀을 찾을 수 없음. 페이지 새로고침 시도...")
                self.driver.refresh()
                time.sleep(5)
                amount_cells = self.driver.find_elements(By.CSS_SELECTOR, "td.td_ri span.fwb")
            
            print(f"      총 {len(amount_cells)}개 금액 셀 발견")
            
            for i, cell in enumerate(amount_cells):
                cell_amount = self._clean_amount(cell.text)
                print(f"      웹 금액 {i+1}: {cell.text} -> {cell_amount}")
                
                if cell_amount == clean_target:
                    print(f"      💡 매칭 성공! 행 {i+1}")
                    
                    # 다양한 체크박스 XPath 시도
                    checkbox_xpaths = [
                        f"/html/body/div[4]/div[3]/div[3]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[3]/div[2]/table/tbody/tr[{i+1}]/td[1]/span/input",
                        f"//table//tr[{i+1}]//input[@name='inp_CardChk']",
                        f"//tr[{i+1}]//input[contains(@name, 'CardChk')]"
                    ]
                    
                    for xpath in checkbox_xpaths:
                        try:
                            print(f"      체크박스 XPath 시도: {xpath}")
                            checkbox = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                            
                            # 스크롤해서 체크박스가 보이도록 하기
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                            time.sleep(1)
                            
                            checkbox.click()
                            time.sleep(2)
                            
                            print(f"      ✅ 체크박스 클릭 성공")
                            return True
                            
                        except Exception as e:
                            print(f"      XPath 실패: {e}")
                            continue
                    
                    print(f"      ❌ 모든 체크박스 XPath 실패")
                    return False
            
            print(f"      ❌ 매칭되는 금액을 찾지 못함")
            return False
            
        except Exception as e:
            print(f"      ❌ 체크박스 클릭 실패: {e}")
            return False

    def _clean_amount(self, amount_text):
        """금액 텍스트 정리"""
        try:
            if not amount_text:
                return "0"
            
            # 모든 특수문자, 공백, 쉼표 제거
            cleaned = str(amount_text).replace(',', '').replace(' ', '').replace('원', '').replace('₩', '')
            
            # 소숫점 처리
            if '.' in cleaned:
                cleaned = cleaned.split('.')[0]
            
            # 숫자가 아닌 문자 제거
            cleaned = re.sub(r'[^\d]', '', cleaned)
            
            # 빈 문자열이면 0
            if not cleaned:
                return "0"
                
            return str(int(cleaned))
        except:
            return "0"

    def _input_form_data(self, data_row):
        """폼 데이터 입력"""
        try:
            # 표준 적요 입력
            if data_row.get('standard_summary'):
                print(f"      표준적요: {data_row['standard_summary']}")
                summary_input = self.driver.find_element(By.ID, "txtExpendCardDispSummary")
                summary_input.clear()
                summary_input.send_keys(data_row['standard_summary'])
                summary_input.send_keys(Keys.TAB)  # ENTER 대신 TAB
                time.sleep(1)
            
            # 증빙 유형 입력
            if data_row.get('evidence_type'):
                print(f"      증빙유형: {data_row['evidence_type']}")
                evidence_input = self.driver.find_element(By.ID, "txtExpendCardDispAuth")
                evidence_input.clear()
                evidence_input.send_keys(data_row['evidence_type'])
                evidence_input.send_keys(Keys.TAB)  # ENTER 대신 TAB
                time.sleep(1)
            
            # 적요 입력
            if data_row.get('note'):
                print(f"      적요: {data_row['note']}")
                note_input = self.driver.find_element(By.ID, "txtExpendCardDispNote")
                note_input.clear()
                note_input.send_keys(data_row['note'])
                note_input.send_keys(Keys.TAB)  # ENTER 대신 TAB
                time.sleep(1)
            
            # 프로젝트 입력
            if data_row.get('project'):
                print(f"      프로젝트: {data_row['project']}")
                project_input = self.driver.find_element(By.ID, "txtExpendCardDispProject")
                project_input.clear()
                project_input.send_keys(data_row['project'])
                # 마지막 필드이므로 TAB 없이
                time.sleep(1)
            
            return True
            
        except Exception as e:
            raise Exception(f"폼 데이터 입력 실패: {e}")

    def _click_save(self):
        """저장 버튼 클릭"""
        try:
            save_btn = self.wait.until(EC.element_to_be_clickable((By.ID, "btnExpendCardInfoSave")))
            save_btn.click()
            time.sleep(3)  # 저장 후 더 긴 대기
            return True
        except Exception as e:
            raise Exception(f"저장 버튼 클릭 실패: {e}")

    def run_automation(self, processed_data, progress_callback=None, user_id="", password=""):
        """메인 자동화 실행 메서드"""
        try:
            print("🚀 자동화 프로세스 시작")
            
            # 1. 브라우저 설정
            if progress_callback:
                progress_callback("브라우저를 시작하는 중...")
            self.setup_driver()
            
            # 2. 로그인
            if progress_callback:
                progress_callback("그룹웨어에 로그인하는 중...")
            self.login_to_groupware(user_id, password)
            
            # 3. 데이터 정보 확인
            total_records = len(processed_data)
            if not processed_data:
                raise Exception("처리할 데이터가 없습니다.")
            
            start_date = processed_data[0].get('start_date', '')
            end_date = processed_data[0].get('end_date', '')
            batch_size = self.config.SAVE_LIMIT  # 100
            
            print(f"📊 총 레코드 수: {total_records}")
            print(f"📅 처리 기간: {start_date} ~ {end_date}")
            print(f"📦 배치 크기: {batch_size}")
            
            # 4. 배치별 처리
            total_batches = (total_records + batch_size - 1) // batch_size
            
            for batch_num in range(total_batches):
                batch_start = batch_num * batch_size
                batch_end = min(batch_start + batch_size, total_records)
                current_batch = processed_data[batch_start:batch_end]
                
                print(f"\n🔄 배치 {batch_num + 1}/{total_batches} 시작 (레코드 {batch_start + 1}~{batch_end})")
                
                if progress_callback:
                    progress_callback(f"배치 {batch_num + 1}/{total_batches} 처리 중...")
                
                # 4-1. 페이지로 이동
                self.navigate_to_expense_page()
                
                # 4-2. 카드 사용내역 설정 (배치당 1회)
                self.setup_card_interface(start_date, end_date)
                
                # 4-3. 각 레코드 처리
                success_count = 0
                for i, data_row in enumerate(current_batch):
                    record_index = batch_start + i + 1
                    
                    if progress_callback:
                        progress_callback(f"레코드 처리 중... ({record_index}/{total_records})")
                    
                    if self.process_single_record(data_row, record_index, total_records):
                        success_count += 1
                
                print(f"✅ 배치 {batch_num + 1} 완료: {success_count}/{len(current_batch)} 성공")
                
                if progress_callback:
                    progress_callback(f"배치 {batch_num + 1} 완료: {success_count}/{len(current_batch)} 성공")
                
                # 배치 간 대기
                if batch_num < total_batches - 1:  # 마지막 배치가 아니면
                    print("⏳ 다음 배치 준비 중...")
                    time.sleep(3)
            
            print("🎉 모든 작업 완료!")
            if progress_callback:
                progress_callback("모든 작업이 완료되었습니다!")
            
        except Exception as e:
            print(f"❌ 자동화 실패: {e}")
            if progress_callback:
                progress_callback(f"작업 중 오류 발생: {str(e)}")
            raise e
        finally:
            if self.driver:
                print("🔚 브라우저 종료")
                self.driver.quit()