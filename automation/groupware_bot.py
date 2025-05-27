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
    """그룹웨어 자동화 클래스"""
    
    def __init__(self):
        self.config = Config()
        self.driver = None
        self.wait = None
        self.save_count = 0
        # 로그인 URL 추가
        self.login_url = "https://gw.meta-m.co.kr/gw/userMain.do"

    def login_to_groupware(self, user_id, password):
        """그룹웨어 로그인"""
        try:
            # 로그인 페이지로 이동
            self.driver.get(self.login_url)
            time.sleep(2)
            
            # 아이디 입력
            id_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "userId"))
            )
            id_input.clear()
            id_input.send_keys(user_id)
            time.sleep(1)
            
            # 비밀번호 입력
            pw_input = self.driver.find_element(By.ID, "userPw")
            pw_input.clear()
            pw_input.send_keys(password)
            time.sleep(1)
            
            # 로그인 버튼 클릭 (Enter 키 사용)
            pw_input.send_keys(Keys.ENTER)
            time.sleep(3)
            
            # 로그인 성공 확인 (URL 변화 또는 특정 요소 확인)
            current_url = self.driver.current_url
            if "userMain.do" not in current_url or "login" in current_url.lower():
                # 로그인 실패 체크 (에러 메시지가 있는지 확인)
                try:
                    error_elements = self.driver.find_elements(By.CLASS_NAME, "error")
                    if error_elements:
                        raise Exception("로그인 정보가 올바르지 않습니다.")
                except:
                    pass
                raise Exception("로그인에 실패했습니다. 아이디와 비밀번호를 확인해주세요.")
            
            return True
            
        except Exception as e:
            raise Exception(f"로그인 실패: {e}")


    def setup_driver(self):
        """WebDriver 설정"""
        options = Options()
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument('--disable-web-security')
        options.add_argument('--allow-running-insecure-content')
        
        try:
            self.driver = webdriver.Chrome(options=options)
            self.wait = WebDriverWait(self.driver, self.config.LOGIN_TIMEOUT)
            return True
        except Exception as e:
            raise Exception(f"브라우저 실행 실패: {e}")
    
    def navigate_to_groupware(self):
        """그룹웨어 사이트로 이동"""
        try:
            self.driver.get(self.config.GROUPWARE_URL)
            time.sleep(3)
            return True
        except Exception as e:
            raise Exception(f"그룹웨어 사이트 이동 실패: {e}")
    
    def click_card_history(self):
        """카드 사용내역 버튼 클릭"""
        try:
            # ID로 먼저 시도
            card_btn = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btnExpendInterfaceCard"))
            )
            card_btn.click()
            time.sleep(2)
            return True
        except:
            # xpath로 재시도
            try:
                card_btn = self.driver.find_element(By.XPATH, self.config.CARD_ELEMENTS["card_history_btn_xpath"])
                card_btn.click()
                time.sleep(2)
                return True
            except Exception as e:
                raise Exception(f"카드 사용내역 버튼 클릭 실패: {e}")
    
    def click_select_button(self):
        """선택 버튼 클릭"""
        try:
            select_btn = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btnExpendCardInfoHelpPop"))
            )
            select_btn.click()
            time.sleep(2)
            return True
        except Exception as e:
            raise Exception(f"선택 버튼 클릭 실패: {e}")
    
    def input_date_range(self, start_date, end_date):
        """날짜 범위 입력 (Kendo UI DatePicker 대응 - 개선된 버전)"""
        try:
            # 날짜 형식 변환 (YYYYMMDD -> YYYY-MM-DD)
            formatted_start = f"{start_date[:4]}-{start_date[4:6]}-{start_date[6:8]}"
            formatted_end = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:8]}"
            
            print(f"입력할 시작날짜: {formatted_start}")
            print(f"입력할 종료날짜: {formatted_end}")
            
            # === 시작 날짜 입력 ===
            start_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "txtExpendCardFromDate"))
            )
            
            # 1. 필드 클릭하여 포커스
            start_input.click()
            time.sleep(1)
            
            # 2. 기존 값 전체 선택 후 삭제
            start_input.send_keys(Keys.CONTROL + "a")
            time.sleep(0.3)
            start_input.send_keys(Keys.DELETE)
            time.sleep(0.3)
            
            # 3. 새 값 입력
            start_input.send_keys(formatted_start)
            time.sleep(1)
            
            # 4. Enter로 값 확정
            start_input.send_keys(Keys.ENTER)
            time.sleep(2)  # 더 긴 대기 시간
            
            print(f"시작날짜 입력 후 값: {start_input.get_attribute('value')}")
            
            # === 종료 날짜 입력 ===
            end_input = self.driver.find_element(By.ID, "txtExpendCardToDate")
            
            # 1. 페이지의 다른 곳 클릭 후 종료 날짜 필드 클릭 (포커스 확실히 이동)
            self.driver.find_element(By.TAG_NAME, "body").click()
            time.sleep(0.5)
            
            end_input.click()
            time.sleep(1)
            
            # 2. 기존 값 전체 선택 후 삭제
            end_input.send_keys(Keys.CONTROL + "a")
            time.sleep(0.3)
            end_input.send_keys(Keys.DELETE)
            time.sleep(0.3)
            
            # 3. 새 값 입력
            end_input.send_keys(formatted_end)
            time.sleep(1)
            
            # 4. Enter로 값 확정
            end_input.send_keys(Keys.ENTER)
            time.sleep(2)
            
            print(f"종료날짜 입력 후 값: {end_input.get_attribute('value')}")
            
            # 5. 값 검증
            actual_start = start_input.get_attribute('value')
            actual_end = end_input.get_attribute('value')
            
            if actual_start != formatted_start or actual_end != formatted_end:
                print(f"⚠️ 날짜 값이 예상과 다릅니다. JavaScript 방법으로 재시도...")
                raise Exception("날짜 값 불일치")
            
            return True
            
        except Exception as e:
            # JavaScript로 직접 값 설정하는 백업 방법
            try:
                print(f"일반 입력 실패, JavaScript로 재시도: {e}")
                
                # JavaScript를 사용하여 강제로 값 설정
                js_script = f"""
                // 시작 날짜 설정
                var startInput = document.getElementById('txtExpendCardFromDate');
                var startDatePicker = $("#txtExpendCardFromDate").data("kendoDatePicker");
                
                if (startDatePicker) {{
                    startDatePicker.value(new Date("{formatted_start}"));
                    startDatePicker.trigger("change");
                }} else {{
                    startInput.value = "{formatted_start}";
                    startInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
                
                // 종료 날짜 설정
                var endInput = document.getElementById('txtExpendCardToDate');
                var endDatePicker = $("#txtExpendCardToDate").data("kendoDatePicker");
                
                if (endDatePicker) {{
                    endDatePicker.value(new Date("{formatted_end}"));
                    endDatePicker.trigger("change");
                }} else {{
                    endInput.value = "{formatted_end}";
                    endInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
                
                return [startInput.value, endInput.value];
                """
                
                result = self.driver.execute_script(js_script)
                time.sleep(2)
                
                print(f"JavaScript 설정 후 값: 시작={result[0]}, 종료={result[1]}")
                
                return True
                
            except Exception as e2:
                # 최후의 수단: 개별 문자 입력
                try:
                    print(f"JavaScript도 실패, 개별 문자 입력 시도: {e2}")
                    
                    # 시작 날짜 개별 입력
                    start_input = self.driver.find_element(By.ID, "txtExpendCardFromDate")
                    start_input.click()
                    time.sleep(0.5)
                    start_input.send_keys(Keys.CONTROL + "a")
                    start_input.send_keys(Keys.DELETE)
                    
                    for char in formatted_start:
                        start_input.send_keys(char)
                        time.sleep(0.1)
                    
                    start_input.send_keys(Keys.TAB)
                    time.sleep(1)
                    
                    # 종료 날짜 개별 입력
                    end_input = self.driver.find_element(By.ID, "txtExpendCardToDate")
                    end_input.click()
                    time.sleep(0.5)
                    end_input.send_keys(Keys.CONTROL + "a")
                    end_input.send_keys(Keys.DELETE)
                    
                    for char in formatted_end:
                        end_input.send_keys(char)
                        time.sleep(0.1)
                    
                    end_input.send_keys(Keys.TAB)
                    time.sleep(1)
                    
                    return True
                    
                except Exception as e3:
                    raise Exception(f"날짜 입력 실패 (모든 방법 실패): {e3}")
            
    def input_date_range_alternative(self, start_date, end_date):
        """대안 방법: 날짜 달력 아이콘 클릭 후 직접 입력"""
        try:
            # 시작 날짜 - 달력 아이콘 클릭
            start_calendar_icon = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#txtExpendCardFromDate + .k-select"))
            )
            start_calendar_icon.click()
            time.sleep(1)
            
            # 달력에서 날짜 선택 로직 (복잡할 수 있음)
            # 또는 input 필드에 직접 입력
            start_input = self.driver.find_element(By.ID, "txtExpendCardFromDate")
            self.driver.execute_script("arguments[0].value = '';", start_input)
            self.driver.execute_script(f"arguments[0].value = '{start_date[:4]}-{start_date[4:6]}-{start_date[6:8]}';", start_input)
            
            # ESC로 달력 닫기
            start_input.send_keys(Keys.ESCAPE)
            time.sleep(1)
            
            # 종료 날짜도 동일한 방식
            end_calendar_icon = self.driver.find_element(By.CSS_SELECTOR, "#txtExpendCardToDate + .k-select")
            end_calendar_icon.click()
            time.sleep(1)
            
            end_input = self.driver.find_element(By.ID, "txtExpendCardToDate")
            self.driver.execute_script("arguments[0].value = '';", end_input)
            self.driver.execute_script(f"arguments[0].value = '{end_date[:4]}-{end_date[4:6]}-{end_date[6:8]}';", end_input)
            
            end_input.send_keys(Keys.ESCAPE)
            time.sleep(1)
            
            return True
            
        except Exception as e:
            raise Exception(f"대안 날짜 입력 실패: {e}")
    def input_date_range_step_by_step(self, start_date, end_date):
        """단계별 디버깅을 위한 날짜 입력 메서드"""
        try:
            formatted_start = f"{start_date[:4]}-{start_date[4:6]}-{start_date[6:8]}"
            formatted_end = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:8]}"
            
            print(f"=== 날짜 입력 시작 ===")
            print(f"원본 시작날짜: {start_date} -> 변환: {formatted_start}")
            print(f"원본 종료날짜: {end_date} -> 변환: {formatted_end}")
            
            # 시작 날짜 처리
            print("1. 시작날짜 필드 찾기...")
            start_input = self.wait.until(EC.presence_of_element_located((By.ID, "txtExpendCardFromDate")))
            print(f"   시작날짜 필드 현재 값: {start_input.get_attribute('value')}")
            
            print("2. 시작날짜 입력...")
            start_input.click()
            time.sleep(1)
            start_input.send_keys(Keys.CONTROL + "a")
            start_input.send_keys(formatted_start)
            start_input.send_keys(Keys.ENTER)
            time.sleep(2)
            print(f"   시작날짜 입력 후 값: {start_input.get_attribute('value')}")
            
            # 종료 날짜 처리
            print("3. 종료날짜 필드 찾기...")
            end_input = self.driver.find_element(By.ID, "txtExpendCardToDate")
            print(f"   종료날짜 필드 현재 값: {end_input.get_attribute('value')}")
            
            print("4. 종료날짜 입력...")
            # 다른 곳 클릭 후 종료날짜 필드로 포커스 이동
            self.driver.find_element(By.TAG_NAME, "body").click()
            time.sleep(0.5)
            
            end_input.click()
            time.sleep(1)
            print(f"   클릭 후 종료날짜 값: {end_input.get_attribute('value')}")
            
            end_input.send_keys(Keys.CONTROL + "a")
            time.sleep(0.5)
            print(f"   전체선택 후 값: {end_input.get_attribute('value')}")
            
            end_input.send_keys(formatted_end)
            time.sleep(1)
            print(f"   입력 후 값: {end_input.get_attribute('value')}")
            
            end_input.send_keys(Keys.ENTER)
            time.sleep(2)
            print(f"   Enter 후 최종 값: {end_input.get_attribute('value')}")
            
            print("=== 날짜 입력 완료 ===")
            
            return True
            
        except Exception as e:
            print(f"단계별 날짜 입력 실패: {e}")
            raise e
    
    def click_search(self):
        """검색 버튼 클릭"""
        try:
            search_btn = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btnExpendCardListSearch"))
            )
            search_btn.click()
            time.sleep(3)  # 검색 결과 로딩 대기
            return True
        except Exception as e:
            raise Exception(f"검색 버튼 클릭 실패: {e}")
    
    def find_matching_amount_row(self, target_amount):
        """금액이 일치하는 행 찾기 (강화된 정수 비교)"""
        try:
            # 타겟 금액을 정수로 정리
            clean_target = self.clean_web_amount(str(target_amount))
            
            print(f"찾는 금액: {clean_target}")  # 디버깅용
            
            # 금액 셀들 찾기
            amount_cells = self.driver.find_elements(By.CSS_SELECTOR, "td.td_ri span.fwb")
            
            for i, cell in enumerate(amount_cells):
                cell_amount = self.clean_web_amount(cell.text)
                print(f"웹 금액 {i+1}: {cell.text} -> {cell_amount}")  # 디버깅용
                
                if cell_amount == clean_target:
                    # 해당 행의 체크박스 찾기
                    row_index = i + 1  # 1부터 시작
                    checkbox_xpath = f"/html/body/div[4]/div[3]/div[3]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[3]/div[2]/table/tbody/tr[{row_index}]/td[1]/span/input"
                    print(f"매칭 성공! 행 인덱스: {row_index}")  # 디버깅용
                    return checkbox_xpath
            
            print(f"매칭되는 금액을 찾지 못했습니다: {clean_target}")  # 디버깅용
            return None
        except Exception as e:
            raise Exception(f"일치하는 금액 행 찾기 실패: {e}")
        

    def clean_web_amount(self, amount_text):
        """웹 페이지 금액 텍스트 정리"""
        try:
            if not amount_text:
                return "0"
            
            # 모든 특수문자, 공백, 쉼표 제거
            cleaned = str(amount_text).replace(',', '').replace(' ', '').replace('원', '').replace('₩', '')
            
            # 소숫점 처리
            if '.' in cleaned:
                cleaned = cleaned.split('.')[0]
            
            # 숫자가 아닌 문자 제거
            import re
            cleaned = re.sub(r'[^\d]', '', cleaned)
            
            # 빈 문자열이면 0
            if not cleaned:
                return "0"
                
            return str(int(cleaned))
        except:
            return "0"
        
    def click_checkbox(self, checkbox_xpath):
        """체크박스 클릭"""
        try:
            checkbox = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, checkbox_xpath))
            )
            checkbox.click()
            time.sleep(1)
            return True
        except Exception as e:
            raise Exception(f"체크박스 클릭 실패: {e}")
    
    def input_form_data(self, data_row):
        """폼 데이터 입력"""
        try:
            elements = self.config.CARD_ELEMENTS
            
            # 표준 적요 입력
            if data_row.get('standard_summary'):
                summary_input = self.driver.find_element(By.ID, "txtExpendCardDispSummary")
                summary_input.clear()
                summary_input.send_keys(data_row['standard_summary'])
                summary_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 증빙 유형 입력
            if data_row.get('evidence_type'):
                evidence_input = self.driver.find_element(By.ID, "txtExpendCardDispAuth")
                evidence_input.clear()
                evidence_input.send_keys(data_row['evidence_type'])
                evidence_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 적요 입력
            if data_row.get('note'):
                note_input = self.driver.find_element(By.ID, "txtExpendCardDispNote")
                note_input.clear()
                note_input.send_keys(data_row['note'])
                note_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 프로젝트 입력
            if data_row.get('project'):
                project_input = self.driver.find_element(By.ID, "txtExpendCardDispProject")
                project_input.clear()
                project_input.send_keys(data_row['project'])
                project_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            return True
        except Exception as e:
            raise Exception(f"폼 데이터 입력 실패: {e}")
    
    def click_save(self):
        """저장 버튼 클릭"""
        try:
            save_btn = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btnExpendCardInfoSave"))
            )
            save_btn.click()
            time.sleep(2)
            self.save_count += 1
            return True
        except Exception as e:
            raise Exception(f"저장 버튼 클릭 실패: {e}")
    
    def process_single_record(self, data_row):
        """단일 레코드 처리 - 올바른 순서로 수정"""
        try:
            # 3. 금액이 일치하는 행 찾기
            checkbox_xpath = self.find_matching_amount_row(data_row.get('amount', ''))
            
            if not checkbox_xpath:
                raise Exception(f"일치하는 금액을 찾을 수 없습니다: {data_row.get('amount', '')}")
            
            # 3. 체크박스 클릭
            self.click_checkbox(checkbox_xpath)
            
            # 4-7. 폼 데이터 입력 (순서대로)
            self.input_form_data_step_by_step(data_row)
            
            # 8. 저장
            self.click_save()
            
            return True
            
        except Exception as e:
            raise Exception(f"레코드 처리 실패: {e}")
        
    def input_form_data_step_by_step(self, data_row):
        """폼 데이터 단계별 입력"""
        try:
            # 4. [표준 적요] 입력
            if data_row.get('standard_summary'):
                summary_input = self.driver.find_element(By.ID, "txtExpendCardDispSummary")
                summary_input.clear()
                summary_input.send_keys(data_row['standard_summary'])
                summary_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 5. [증빙 유형] 입력
            if data_row.get('evidence_type'):
                evidence_input = self.driver.find_element(By.ID, "txtExpendCardDispAuth")
                evidence_input.clear()
                evidence_input.send_keys(data_row['evidence_type'])
                evidence_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 6. [적요] 입력
            if data_row.get('note'):
                note_input = self.driver.find_element(By.ID, "txtExpendCardDispNote")
                note_input.clear()
                note_input.send_keys(data_row['note'])
                note_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 7. [프로젝트] 입력
            if data_row.get('project'):
                project_input = self.driver.find_element(By.ID, "txtExpendCardDispProject")
                project_input.clear()
                project_input.send_keys(data_row['project'])
                project_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            return True
        except Exception as e:
            raise Exception(f"폼 데이터 입력 실패: {e}")
    
    def run_automation(self, processed_data, progress_callback=None, user_id="", password=""):
        """자동화 실행 - 올바른 프로세스 순서"""
        try:
            # WebDriver 설정
            if progress_callback:
                progress_callback("브라우저를 시작하는 중...")
            self.setup_driver()
            
            # 로그인
            if progress_callback:
                progress_callback("그룹웨어에 로그인하는 중...")
            self.login_to_groupware(user_id, password)
            
            # 날짜 범위 설정 (첫 번째 데이터에서 추출)
            if processed_data:
                start_date = processed_data[0].get('start_date', '')
                end_date = processed_data[0].get('end_date', '')
            
            total_records = len(processed_data)
            batch_size = self.config.SAVE_LIMIT  # 100
            
            # 배치별로 처리
            for batch_start in range(0, total_records, batch_size):
                batch_end = min(batch_start + batch_size, total_records)
                current_batch = processed_data[batch_start:batch_end]
                
                batch_num = (batch_start // batch_size) + 1
                if progress_callback:
                    progress_callback(f"배치 {batch_num} 시작 ({len(current_batch)}개 레코드)")
                
                # === 1. 그룹웨어 지출결의서 페이지로 이동 ===
                if progress_callback:
                    progress_callback(f"배치 {batch_num}: 지출결의서 페이지로 이동 중...")
                self.navigate_to_groupware()
                
                # === 2. 카드 사용내역 프로세스 시작 ===
                if progress_callback:
                    progress_callback(f"배치 {batch_num}: 카드 사용내역 설정 중...")
                
                # 2-1. [카드 사용내역] 클릭
                self.click_card_history()
                
                # 2-2. [선택] 클릭
                self.click_select_button()
                
                # 2-3. [시작 날짜] 입력
                # 2-4. [종료 날짜] 입력
                self.input_date_range(start_date, end_date)
                
                # 2-5. [검색] 클릭
                self.click_search()
                
                if progress_callback:
                    progress_callback(f"배치 {batch_num}: 검색 완료, 데이터 입력 시작...")
                
                # === 3-8. 각 레코드별 처리 ===
                for i, data_row in enumerate(current_batch):
                    current_index = batch_start + i + 1
                    if progress_callback:
                        progress_callback(f"레코드 처리 중... ({current_index}/{total_records})")
                    
                    try:
                        # 3. 금액 매칭하여 체크박스 클릭
                        checkbox_xpath = self.find_matching_amount_row(data_row.get('amount', ''))
                        
                        if checkbox_xpath:
                            self.click_checkbox(checkbox_xpath)
                            
                            # 4-7. 폼 데이터 입력
                            self.input_form_data(data_row)
                            
                            # 8. 저장
                            self.click_save()
                            
                            if progress_callback:
                                progress_callback(f"레코드 저장 완료 ({current_index}/{total_records})")
                        else:
                            if progress_callback:
                                progress_callback(f"⚠️ 금액 매칭 실패: {data_row.get('amount', '')} (레코드 {current_index})")
                            continue
                            
                    except Exception as e:
                        if progress_callback:
                            progress_callback(f"⚠️ 레코드 {current_index} 처리 실패: {str(e)}")
                        continue
                
                if progress_callback:
                    progress_callback(f"배치 {batch_num} 완료 ({len(current_batch)}개 처리)")
                
                # 배치 완료 후 잠시 대기
                time.sleep(2)
            
            if progress_callback:
                progress_callback("모든 작업이 완료되었습니다!")
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"작업 중 오류 발생: {str(e)}")
            raise e
        finally:
            if self.driver:
                self.driver.quit()