import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
import threading
from automation.groupware_bot import GroupwareAutomation
from automation.data_processor import DataProcessor

class ExpenseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("지출경비서 자동입력 프로그램 v1.0")
        self.root.geometry("500x600")  
        self.root.resizable(True, True)  
        
        self.center_window()
        self.init_variables()
        self.create_widgets()
        
        # 자동화 객체 초기화
        self.automation = GroupwareAutomation()
        self.data_processor = DataProcessor()
        
    def center_window(self):
        """창을 화면 중앙에 위치"""
        self.root.update_idletasks()
        width = 500
        height = 600 
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def init_variables(self):
        """변수 초기화"""
        self.file_path = ""
        self.category = tk.StringVar(value="해외결제 법인카드")
        self.start_date = tk.StringVar()
        self.end_date = tk.StringVar()
        self.progress_var = tk.StringVar(value="대기 중...")
        # 로그인 정보 추가
        self.user_id = tk.StringVar()
        self.user_password = tk.StringVar()
    
    def create_widgets(self):
        """GUI 위젯 생성"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(main_frame, text="지출경비서 자동입력", 
                            font=("맑은 고딕", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 1. 로그인 정보
        self.create_login_section(main_frame)
        
        # 2. 파일 업로드
        self.create_file_section(main_frame)
        
        # 3. 카테고리 설정
        self.create_category_section(main_frame)
        
        # 4. 날짜 설정
        self.create_date_section(main_frame)
        
        # 5. 실행 버튼 및 진행상황
        self.create_control_section(main_frame)
    
    def create_login_section(self, parent):
        """로그인 정보 섹션"""
        login_frame = ttk.LabelFrame(parent, text="1. 로그인 정보", padding="8")
        login_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 아이디
        id_frame = ttk.Frame(login_frame)
        id_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(id_frame, text="아이디:", width=8).pack(side=tk.LEFT)
        ttk.Entry(id_frame, textvariable=self.user_id, width=20).pack(side=tk.LEFT, padx=(5, 0))
        
        # 비밀번호
        pw_frame = ttk.Frame(login_frame)
        pw_frame.pack(fill=tk.X)
        ttk.Label(pw_frame, text="비밀번호:", width=8).pack(side=tk.LEFT)
        ttk.Entry(pw_frame, textvariable=self.user_password, show="*", width=20).pack(side=tk.LEFT, padx=(5, 0))


    def create_file_section(self, parent):
        """파일 업로드 섹션"""
        file_frame = ttk.LabelFrame(parent, text="2. 파일 업로드", padding="8")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="파일을 선택해주세요", 
                                foreground="gray")
        self.file_label.pack(anchor=tk.W, pady=(0, 8))
        
        file_button = ttk.Button(file_frame, text="📁 CSV/Excel 파일 선택", 
                                command=self.select_file)
        file_button.pack()
    
    def create_category_section(self, parent):
        """카테고리 설정 섹션"""
        category_frame = ttk.LabelFrame(parent, text="3. 카테고리 설정", padding="8")
        category_frame.pack(fill=tk.X, pady=(0, 10))
        
        categories = ["해외결제 법인카드", "그 외"]
        for category in categories:
            radio = ttk.Radiobutton(category_frame, text=category, 
                                variable=self.category, value=category)
            radio.pack(anchor=tk.W, pady=2)
    
    def create_date_section(self, parent):
        """날짜 설정 섹션"""
        date_frame = ttk.LabelFrame(parent, text="4. 날짜 설정", padding="8")
        date_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 시작날짜
        start_frame = ttk.Frame(date_frame)
        start_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(start_frame, text="시작날짜:", width=8).pack(side=tk.LEFT)
        ttk.Entry(start_frame, textvariable=self.start_date, width=12).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(start_frame, text="(예: 20250501)", foreground="gray", font=("맑은 고딕", 8)).pack(side=tk.LEFT)
        
        # 종료날짜
        end_frame = ttk.Frame(date_frame)
        end_frame.pack(fill=tk.X)
        ttk.Label(end_frame, text="종료날짜:", width=8).pack(side=tk.LEFT)
        ttk.Entry(end_frame, textvariable=self.end_date, width=12).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(end_frame, text="(예: 20250531)", foreground="gray", font=("맑은 고딕", 8)).pack(side=tk.LEFT)

        
    def create_control_section(self, parent):
        """실행 버튼 및 진행상황 섹션"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=20)
        
        self.start_button = ttk.Button(button_frame, text="🚀 지출결의서 작업 시작", 
                                      command=self.start_automation_thread)
        self.start_button.pack(pady=10)
        
        self.progress_label = ttk.Label(parent, textvariable=self.progress_var)
        self.progress_label.pack(pady=(10, 0))
        
        self.progress_bar = ttk.Progressbar(parent, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
    
    def select_file(self):
        """파일 선택"""
        file_types = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]
        
        file_path = filedialog.askopenfilename(
            title="파일을 선택해주세요",
            filetypes=file_types
        )
        
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"✅ {filename}", foreground="green")
    
    def validate_inputs(self):
        """입력값 검증"""
        if not self.user_id.get() or not self.user_password.get():
            messagebox.showerror("오류", "아이디와 비밀번호를 입력해주세요!")
            return False
        
        if not self.file_path:
            messagebox.showerror("오류", "파일을 선택해주세요!")
            return False
        
        if not self.start_date.get() or not self.end_date.get():
            messagebox.showerror("오류", "날짜를 입력해주세요!")
            return False
        
        try:
            datetime.strptime(self.start_date.get(), "%Y%m%d")
            datetime.strptime(self.end_date.get(), "%Y%m%d")
        except ValueError:
            messagebox.showerror("오류", "날짜 형식이 올바르지 않습니다! (예: 20250501)")
            return False
        
        return True
    
    def start_automation_thread(self):
        """별도 스레드에서 자동화 실행"""
        if not self.validate_inputs():
            return
        
        # UI 상태 변경
        self.start_button.config(state="disabled")
        self.progress_bar.start()
        
        # 별도 스레드에서 실행 (UI 블로킹 방지)
        thread = threading.Thread(target=self.run_automation)
        thread.daemon = True
        thread.start()
    
    def run_automation(self):
        """자동화 실행"""
        try:
            # 1. 데이터 처리
            self.update_progress("파일을 읽는 중...")
            data = self.data_processor.load_file(self.file_path)
            
            self.update_progress("데이터를 처리하는 중...")
            processed_data = self.data_processor.process_data(
                data, self.category.get(), 
                self.start_date.get(), self.end_date.get()
            )
            
            # 2. 그룹웨어 자동화 실행 (로그인 정보 포함)
            self.update_progress("그룹웨어에 접속하는 중...")
            self.automation.run_automation(
                processed_data, 
                progress_callback=self.update_progress,
                user_id=self.user_id.get(),
                password=self.user_password.get()
            )
            
            # 완료
            self.update_progress("작업이 완료되었습니다!")
            self.show_completion_message()
            
        except Exception as e:
            self.update_progress("오류가 발생했습니다.")
            self.show_error_message(str(e))
        
        finally:
            self.reset_ui()
    
    def update_progress(self, message):
        """진행상황 업데이트 (스레드 안전)"""
        self.root.after(0, lambda: self.progress_var.set(message))
    
    def show_completion_message(self):
        """완료 메시지 표시 (스레드 안전)"""
        self.root.after(0, lambda: messagebox.showinfo("완료", "지출결의서 입력이 완료되었습니다!"))
    
    def show_error_message(self, error):
        """오류 메시지 표시 (스레드 안전)"""
        self.root.after(0, lambda: messagebox.showerror("오류", f"작업 중 오류가 발생했습니다:\n{error}"))
    
    def reset_ui(self):
        """UI 상태 초기화 (스레드 안전)"""
        def reset():
            self.progress_bar.stop()
            self.start_button.config(state="normal")
        
        self.root.after(0, reset)