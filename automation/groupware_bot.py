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
            
            print(f"📊 총 레코드 수: {total_records}")
            print(f"📅 처리 기간: {start_date} ~ {end_date}")
            
            # 4. 첫 번째 페이지 이동 (최초 1회만)
            self.navigate_to_expense_page()
            
            # 5. 데이터 처리 루프
            processed_count = 0
            round_number = 1
            
            while processed_count < total_records:
                print(f"\n🔄 처리 라운드 {round_number} 시작 (진행률: {processed_count}/{total_records})")
                
                if progress_callback:
                    progress_callback(f"라운드 {round_number} 처리 중... ({processed_count}/{total_records})")
                
                # 5-1. 카드 사용내역 설정 (페이지 이동 없이)
                self.setup_card_interface(start_date, end_date)
                
                # 5-2. 현재 페이지에서 처리 가능한 모든 데이터 입력
                round_processed = 0
                
                for i in range(processed_count, total_records):
                    data_row = processed_data[i]
                    record_index = i + 1
                    
                    if progress_callback:
                        progress_callback(f"레코드 처리 중... ({record_index}/{total_records})")
                    
                    # 개별 레코드 처리 (체크박스 → 입력 → 저장)
                    if self.process_single_record(data_row, record_index, total_records):
                        round_processed += 1
                        processed_count += 1
                    else:
                        # 현재 페이지에서 더 이상 처리할 수 없으면 중단
                        print(f"   💡 현재 페이지에서 더 이상 처리할 데이터가 없음")
                        break
                
                print(f"✅ 라운드 {round_number} 완료: {round_processed}개 처리됨")
                
                # 5-3. 처리된 데이터가 있으면 전체 체크박스 클릭 후 반영
                if round_processed > 0:
                    print("🔄 전체 체크박스 클릭 및 반영 시작...")
                    
                    # 전체 체크박스 클릭
                    if self._click_select_all_checkbox():
                        # 반영 버튼 클릭 및 완료 대기
                        if self._click_apply_button():
                            print(f"✅ {round_processed}개 데이터 반영 완료")
                            print("📋 반영된 데이터는 누적되었으며, 같은 화면에서 계속 진행합니다")
                            time.sleep(2)  # 반영 후 안정화 대기
                        else:
                            print("❌ 반영 실패")
                            break
                    else:
                        print("❌ 전체 체크박스 클릭 실패")
                        break
                else:
                    # 더 이상 처리할 데이터가 없으면 종료
                    print("🔚 모든 데이터 처리 완료")
                    break
                
                round_number += 1
            
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
        
    def _click_apply_button(self):
        """반영 버튼 클릭 및 완료까지 대기"""
        try:
            print("🔄 반영 버튼 클릭...")
            
            # config에서 반영 버튼 정보 가져오기
            apply_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, self.config.CARD_ELEMENTS["apply_btn"])))
            apply_btn.click()
            print("✅ 반영 버튼 클릭 완료")
            
            # 반영 진행률 팝업이 나타날 때까지 대기
            time.sleep(2)
            
            # 반영 완료까지 대기
            if self._wait_for_apply_completion():
                print("✅ 반영 완료")
                return True
            else:
                print("❌ 반영 대기 중 오류 발생")
                return False
            
        except Exception as e:
            print(f"❌ 반영 버튼 클릭 실패: {e}")
            # 백업 방법
            try:
                apply_btn = self.driver.find_element(By.ID, "btnExpendCardToExpend")
                apply_btn.click()
                print("✅ 반영 버튼 클릭 완료 (백업 방법)")
                
                # 반영 완료까지 대기
                if self._wait_for_apply_completion():
                    print("✅ 반영 완료")
                    return True
                else:
                    return False
            except:
                return False    
    
    def _wait_for_apply_completion(self):
        """반영 진행률 팝업이 사라질 때까지 대기"""
        try:
            print("⏳ 반영 진행률 팝업 대기 중...")
            
            # 팝업이 나타날 때까지 잠시 대기
            time.sleep(3)
            
            # 팝업 선택자들
            popup_selectors = [
                (By.ID, "PLP_divMainProgPop"),
                (By.CSS_SELECTOR, "div[id='PLP_divMainProgPop']"),
                (By.XPATH, "//div[@id='PLP_divMainProgPop']")
            ]
            
            # 팝업이 나타났는지 확인
            popup_appeared = False
            for selector_type, selector in popup_selectors:
                try:
                    popup = self.driver.find_element(selector_type, selector)
                    if popup.is_displayed():
                        popup_appeared = True
                        print("📊 반영 진행률 팝업 감지됨")
                        break
                except:
                    continue
            
            if not popup_appeared:
                print("💡 반영 팝업이 감지되지 않음 - 즉시 완료된 것으로 판단")
                time.sleep(5)  # 안전을 위한 추가 대기
                return True
            
            # 팝업이 사라질 때까지 대기
            max_wait_time = 300  # 최대 5분 대기
            wait_count = 0
            
            while wait_count < max_wait_time:
                try:
                    # 진행률 확인
                    progress_element = self.driver.find_element(By.ID, "PLP_txtProgValue")
                    progress_text = progress_element.text.strip()
                    print(f"📈 반영 진행률: {progress_text}")
                    
                    # 총 건수와 실패 건수 확인
                    try:
                        total_count = self.driver.find_element(By.ID, "PLP_txtFullCnt").text.strip()
                        error_count = self.driver.find_element(By.ID, "PLP_txtErrorCnt").text.strip()
                        print(f"📋 총 {total_count}건 (실패 {error_count}건)")
                    except:
                        pass
                    
                    # 팝업이 여전히 존재하는지 확인
                    popup_still_exists = False
                    for selector_type, selector in popup_selectors:
                        try:
                            popup = self.driver.find_element(selector_type, selector)
                            if popup.is_displayed():
                                popup_still_exists = True
                                break
                        except:
                            continue
                    
                    if not popup_still_exists:
                        print("✅ 반영 팝업이 사라짐 - 반영 완료!")
                        time.sleep(2)  # 페이지 전환 안정화 대기
                        return True
                    
                    time.sleep(2)  # 2초마다 확인
                    wait_count += 2
                    
                except Exception as e:
                    # 팝업이 사라졌을 가능성
                    print(f"💡 팝업 요소 접근 실패 (사라진 것으로 판단): {e}")
                    time.sleep(2)
                    return True
            
            print("⚠️ 반영 대기 시간 초과 (5분)")
            return False
            
        except Exception as e:
            print(f"❌ 반영 완료 대기 실패: {e}")
            return False

    def _click_select_all_checkbox(self):
        """전체 체크박스 클릭하여 모든 항목 선택"""
        try:
            print("🔄 전체 체크박스 클릭...")
            
            # 전체 체크박스 찾기 - 여러 방법 시도
            select_all_selectors = [
                (By.XPATH, "/html/body/div[4]/div[3]/div[3]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[3]/div[1]/div/table/thead/tr/th[1]/input"),
                (By.ID, "inp_ListChk"),
                (By.NAME, "inp_ListChk"),
                (By.CSS_SELECTOR, "input[id='inp_ListChk']")
            ]
            
            select_all_btn = None
            for selector_type, selector in select_all_selectors:
                try:
                    select_all_btn = self.wait.until(EC.element_to_be_clickable((selector_type, selector)))
                    break
                except:
                    continue
            
            if not select_all_btn:
                print("❌ 전체 체크박스를 찾을 수 없음")
                return False
            
            select_all_btn.click()
            time.sleep(2)  # 체크박스 선택 처리 대기
            
            print("✅ 전체 체크박스 클릭 완료")
            return True
            
        except Exception as e:
            print(f"❌ 전체 체크박스 클릭 실패: {e}")
            return False