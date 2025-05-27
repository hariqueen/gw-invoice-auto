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
    """그룹웨어 자동화 클래스 - 완전히 새로 작성"""
    
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
            time.sleep(5)
            
            print("✅ 카드 사용내역 설정 완료")
            return True
            
        except Exception as e:
            raise Exception(f"카드 사용내역 설정 실패: {e}")

    def _input_dates(self, start_date, end_date):
        """날짜 입력 (내부 메서드)"""
        try:
            # 날짜 형식 변환 (YYYYMMDD -> YYYY-MM-DD)
            formatted_start = f"{start_date[:4]}-{start_date[4:6]}-{start_date[6:8]}"
            formatted_end = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:8]}"
            
            print(f"    시작날짜 입력: {formatted_start}")
            
            # 시작 날짜 입력
            start_input = self.wait.until(EC.presence_of_element_located((By.ID, "txtExpendCardFromDate")))
            self._clear_and_input(start_input, formatted_start)
            time.sleep(2)
            
            print(f"    종료날짜 입력: {formatted_end}")
            
            # 종료 날짜 입력
            end_input = self.driver.find_element(By.ID, "txtExpendCardToDate")
            self._clear_and_input(end_input, formatted_end)
            time.sleep(2)
            
            # 검증
            actual_start = start_input.get_attribute('value')
            actual_end = end_input.get_attribute('value')
            print(f"    입력 확인 - 시작: {actual_start}, 종료: {actual_end}")
            
            return True
            
        except Exception as e:
            # JavaScript 백업 방법
            print(f"    키보드 입력 실패, JavaScript로 재시도: {e}")
            return self._input_dates_with_javascript(formatted_start, formatted_end)

    def _clear_and_input(self, element, value):
        """요소 클리어 후 값 입력"""
        element.click()
        time.sleep(0.5)
        element.send_keys(Keys.CONTROL + "a")
        time.sleep(0.3)
        element.send_keys(Keys.DELETE)
        time.sleep(0.3)
        element.send_keys(value)
        time.sleep(0.5)
        element.send_keys(Keys.ENTER)
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
        """금액 매칭하여 체크박스 클릭 - 이미 처리된 항목 건너뛰기"""
        try:
            clean_target = self._clean_amount(str(target_amount))
            print(f"      찾는 금액: {clean_target}")
            
            # 금액 셀들 찾기
            amount_cells = self.driver.find_elements(By.CSS_SELECTOR, "td.td_ri span.fwb")
            print(f"      총 {len(amount_cells)}개 금액 셀 발견")
            
            for i, cell in enumerate(amount_cells):
                cell_amount = self._clean_amount(cell.text)
                print(f"      웹 금액 {i+1}: {cell.text} -> {cell_amount}")
                
                if cell_amount == clean_target:
                    print(f"      💡 금액 매칭! 행 {i+1}")
                    
                    # 이미 처리된 행인지 확인
                    if self._is_row_already_processed(i):
                        print(f"      ⏭️ 행 {i+1}은 이미 처리됨 - 건너뛰기")
                        continue  # 다음 매칭되는 행으로 넘어감
                    
                    print(f"      🎯 행 {i+1}을 처리합니다")
                    
                    # 기존 체크박스 클릭 로직 그대로 사용
                    row_index = i + 1
                    
                    # 방법 1: label 클릭
                    print(f"      🔄 방법1 시도: label 클릭")
                    checkbox_label_xpath = f"/html/body/div[4]/div[3]/div[3]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[3]/div[2]/table/tbody/tr[{row_index}]/td[1]/span/label"
                    try:
                        checkbox_label = self.wait.until(EC.element_to_be_clickable((By.XPATH, checkbox_label_xpath)))
                        checkbox_label.click()
                        time.sleep(1)
                        print(f"      ✅ 성공! 방법1(label 클릭)으로 체크박스 클릭 완료")
                        return True
                    except Exception as e:
                        print(f"      ❌ 방법1 실패: {e}")
                    
                    # 방법 2: input 클릭
                    print(f"      🔄 방법2 시도: input 클릭")
                    checkbox_input_xpath = f"/html/body/div[4]/div[3]/div[3]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[3]/div[2]/table/tbody/tr[{row_index}]/td[1]/span/input"
                    try:
                        checkbox_input = self.wait.until(EC.element_to_be_clickable((By.XPATH, checkbox_input_xpath)))
                        checkbox_input.click()
                        time.sleep(1)
                        print(f"      ✅ 성공! 방법2(input 클릭)으로 체크박스 클릭 완료")
                        return True
                    except Exception as e:
                        print(f"      ❌ 방법2 실패: {e}")
                    
                    # 방법 3: name으로 찾기
                    print(f"      🔄 방법3 시도: name 속성으로 찾기")
                    try:
                        checkboxes = self.driver.find_elements(By.NAME, "inp_CardChk")
                        print(f"      name='inp_CardChk'로 {len(checkboxes)}개 체크박스 발견")
                        if i < len(checkboxes):
                            checkboxes[i].click()
                            time.sleep(1)
                            print(f"      ✅ 성공! 방법3(name 속성)으로 체크박스 클릭 완료")
                            return True
                        else:
                            print(f"      ❌ 방법3 실패: 인덱스 {i}가 체크박스 개수 {len(checkboxes)}를 초과")
                    except Exception as e:
                        print(f"      ❌ 방법3 실패: {e}")
                    
                    print(f"      ❌ 모든 방법 실패")
                    return False
            
            print(f"      ❌ 매칭되는 금액을 찾지 못함")
            return False
            
        except Exception as e:
            print(f"      ❌ 체크박스 클릭 실패: {e}")
            return False

    def _is_row_already_processed(self, row_index):
        """해당 행이 이미 처리되었는지 확인 - 마지막 컬럼의 span 태그 확인"""
        try:
            # 해당 행의 마지막 td (4번째 컬럼) 확인
            row_xpath = f"/html/body/div[4]/div[3]/div[3]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[3]/div[2]/table/tbody/tr[{row_index + 1}]"
            
            try:
                row_element = self.driver.find_element(By.XPATH, row_xpath)
                # 마지막 td 찾기 (4번째 컬럼)
                last_td = row_element.find_element(By.CSS_SELECTOR, "td:last-child")
                
                # td 내부의 모든 span 태그 확인
                spans = last_td.find_elements(By.TAG_NAME, "span")
                
                # span이 없거나 모든 span이 비어있으면 미처리
                if not spans:
                    print(f"        ✅ 행 {row_index+1}은 아직 처리되지 않음 (span 없음)")
                    return False
                
                # span들에 의미있는 데이터가 있는지 확인
                has_data = False
                span_contents = []
                
                for span in spans:
                    span_text = span.text.strip()
                    span_id = span.get_attribute("id")
                    
                    # span에 텍스트나 id가 있으면 데이터가 있는 것
                    if span_text or span_id:
                        has_data = True
                        span_contents.append(f"'{span_text}'" if span_text else f"id='{span_id}'")
                
                if has_data:
                    print(f"        💡 행 {row_index+1}은 이미 처리됨 (span 데이터: {', '.join(span_contents)})")
                    return True
                else:
                    print(f"        ✅ 행 {row_index+1}은 아직 처리되지 않음 (span들이 모두 비어있음)")
                    return False
                    
            except Exception as e:
                print(f"        ❓ 행 {row_index+1} 확인 실패: {e} - 미처리로 간주")
                return False
            
        except Exception as e:
            print(f"        ❓ 처리 여부 확인 실패: {e} - 미처리로 간주")
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
                summary_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 증빙 유형 입력
            if data_row.get('evidence_type'):
                print(f"      증빙유형: {data_row['evidence_type']}")
                evidence_input = self.driver.find_element(By.ID, "txtExpendCardDispAuth")
                evidence_input.clear()
                evidence_input.send_keys(data_row['evidence_type'])
                evidence_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 적요 입력
            if data_row.get('note'):
                print(f"      적요: {data_row['note']}")
                note_input = self.driver.find_element(By.ID, "txtExpendCardDispNote")
                note_input.clear()
                note_input.send_keys(data_row['note'])
                note_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            # 프로젝트 입력
            if data_row.get('project'):
                print(f"      프로젝트: {data_row['project']}")
                project_input = self.driver.find_element(By.ID, "txtExpendCardDispProject")
                project_input.clear()
                project_input.send_keys(data_row['project'])
                project_input.send_keys(Keys.ENTER)
                time.sleep(1)
            
            return True
            
        except Exception as e:
            raise Exception(f"폼 데이터 입력 실패: {e}")

    def _click_save(self):
        """저장 버튼 클릭"""
        try:
            save_btn = self.wait.until(EC.element_to_be_clickable((By.ID, "btnExpendCardInfoSave")))
            save_btn.click()
            time.sleep(2)
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
                
                # 4-3. 모든 페이지에서 레코드 처리
                success_count = self.process_all_pages_in_batch(
                    current_batch, batch_start, total_records, progress_callback
                )
                
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


    def process_all_pages_in_batch(self, current_batch, batch_start, total_records, progress_callback=None):
        """배치 내 모든 페이지 처리"""
        try:
            total_processed = 0
            page_num = 1
            
            while True:
                print(f"\n📄 페이지 {page_num} 처리 시작")
                
                # 현재 페이지에서 처리할 수 있는 데이터 찾기
                page_processed = 0
                
                for i, data_row in enumerate(current_batch):
                    record_index = batch_start + total_processed + 1
                    
                    # 이미 처리된 데이터는 건너뛰기
                    if total_processed >= len(current_batch):
                        break
                    
                    if progress_callback:
                        progress_callback(f"레코드 처리 중... ({record_index}/{total_records})")
                    
                    if self.process_single_record(data_row, record_index, total_records):
                        page_processed += 1
                        total_processed += 1
                        
                        # 현재 배치의 모든 데이터를 처리했으면 종료
                        if total_processed >= len(current_batch):
                            print(f"✅ 배치 내 모든 데이터 처리 완료")
                            break
                    else:
                        # 현재 페이지에서 처리할 데이터가 없으면 다음 페이지로
                        break
                
                print(f"📄 페이지 {page_num} 완료: {page_processed}개 처리됨")
                
                # 현재 배치의 모든 데이터를 처리했으면 종료
                if total_processed >= len(current_batch):
                    break
                
                # 다음 페이지로 이동 시도
                if not self._go_to_next_page():
                    print("🔚 더 이상 다음 페이지가 없거나 이동 실패")
                    break
                    
                page_num += 1
                time.sleep(2)  # 페이지 로딩 대기
            
            print(f"✅ 총 {total_processed}개 레코드 처리 완료")
            
            # 모든 처리가 끝나면 반영 버튼 클릭
            self._click_apply_button()
            
            return total_processed
            
        except Exception as e:
            print(f"❌ 페이지 처리 실패: {e}")
            return total_processed
        
    def _go_to_next_page(self):
        """다음 페이지로 이동"""
        try:
            print("🔄 다음 페이지로 이동 시도...")
            
            # 다음 버튼 찾기 - 여러 방법 시도
            next_selectors = [
                (By.ID, "tblExpendCardList_next"),
                (By.XPATH, "/html/body/div[4]/div[3]/div[3]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[5]/a[2]"),
                (By.CSS_SELECTOR, "a.paginate_button.next"),
                (By.XPATH, "//a[contains(@class, 'paginate_button') and contains(@class, 'next')]")
            ]
            
            next_btn = None
            for selector_type, selector in next_selectors:
                try:
                    next_btn = self.driver.find_element(selector_type, selector)
                    break
                except:
                    continue
            
            if not next_btn:
                print("❌ 다음 버튼을 찾을 수 없음")
                return False
            
            # 다음 버튼이 활성화되어 있는지 확인
            btn_class = next_btn.get_attribute("class")
            if "disabled" in btn_class:
                print("🔚 다음 버튼이 비활성화됨 - 마지막 페이지")
                return False
            
            # 다음 버튼 클릭
            next_btn.click()
            time.sleep(3)  # 페이지 로딩 대기
            
            print("✅ 다음 페이지로 이동 완료")
            return True
            
        except Exception as e:
            print(f"❌ 다음 페이지 이동 실패: {e}")
            return False
    
    def _click_apply_button(self):
        """반영 버튼 클릭"""
        try:
            print("🔄 반영 버튼 클릭...")
            
            # config에서 반영 버튼 정보 가져오기
            apply_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, self.config.CARD_ELEMENTS["apply_btn"])))
            apply_btn.click()
            time.sleep(3)  # 반영 처리 대기
            
            print("✅ 반영 완료")
            return True
            
        except Exception as e:
            print(f"❌ 반영 버튼 클릭 실패: {e}")
            # 백업 방법
            try:
                apply_btn = self.driver.find_element(By.ID, "btnExpendCardToExpend")
                apply_btn.click()
                time.sleep(3)
                print("✅ 반영 완료 (백업 방법)")
                return True
            except:
                return False