"""
PyInstaller로 실행파일 생성 (불필요한 라이브러리 제외)
"""

import os
import subprocess

def build_exe():
    """실행파일 생성 (최적화된 버전)"""
    
    # 제외할 모듈들 (용량 최적화)
    excludes = [
        'torch', 'torchvision', 'torchaudio',  # PyTorch 관련
        'scipy', 'matplotlib', 'sympy',        # 과학 계산 라이브러리
        'PIL',                                 # 이미지/수치 계산
        'tensorflow', 'keras',                 # AI 라이브러리
        'jupyter', 'IPython'                   # 개발 도구
    ]
    
    exclude_args = []
    for module in excludes:
        exclude_args.extend(['--exclude-module', module])
    
    cmd = [
        "pyinstaller",
        "--onedir",                     # 폴더로 생성 (macOS 권장)
        "--windowed",                   # 콘솔창 숨기기
        "--name=지출경비서자동입력",      # 실행파일명
        "--add-data=automation:automation",  # automation 폴더 포함
        "--add-data=gui:gui",               # gui 폴더 포함
        # "--icon=icon.ico",              # 아이콘 
        *exclude_args,                  # 제외 모듈들
        "main.py"
    ]
    
    print("🚀 실행파일 생성을 시작합니다...")
    print("⚠️  불필요한 라이브러리를 제외하여 용량을 최적화합니다.")
    
    try:
        subprocess.run(cmd, check=True)
        print("✅ 실행파일이 dist/지출경비서자동입력/ 폴더에 생성되었습니다!")
        print("📁 실행방법: dist/지출경비서자동입력/지출경비서자동입력 파일을 더블클릭")
    except subprocess.CalledProcessError as e:
        print(f"❌ 빌드 실패: {e}")
    except FileNotFoundError:
        print("❌ PyInstaller가 설치되어 있지 않습니다.")
        print("💡 설치 명령어: pip install pyinstaller")

def create_simple_requirements():
    """필수 라이브러리만 포함한 requirements.txt 생성"""
    requirements = """selenium
pandas
openpyxl
webdriver-manager
pyinstaller
"""
    
    with open('requirements_minimal.txt', 'w', encoding='utf-8') as f:
        f.write(requirements)
    
    print("requirements_minimal.txt 파일을 생성했습니다.")
    print("새 환경에서는 이 파일로 설치하세요: pip install -r requirements_minimal.txt")

if __name__ == "__main__":
    create_simple_requirements()
    build_exe()