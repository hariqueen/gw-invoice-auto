import pandas as pd
from datetime import datetime
from .config import Config

class DataProcessor:
    """데이터 처리 클래스"""
    
    def __init__(self):
        self.config = Config()
    
    def load_file(self, file_path):
        """파일 로드"""
        try:
            if file_path.endswith('.csv'):
                return pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                return pd.read_excel(file_path)
        except Exception as e:
            raise Exception(f"파일을 읽을 수 없습니다: {e}")
    
    def process_data(self, data, category, start_date, end_date):
        """데이터 처리 및 필터링"""
        try:
            # 날짜 형식 변환
            start_dt = datetime.strptime(start_date, "%Y%m%d")
            end_dt = datetime.strptime(end_date, "%Y%m%d")
            
            # 데이터 전처리
            processed_data = []
            
            for index, row in data.iterrows():
                # 실제 컬럼명에 맞게 매핑
                processed_row = {
                    'category': category,
                    'amount': self.clean_amount(row.get('매출금액', 0)),
                    'standard_summary': row.get('표준적요', ''),
                    'evidence_type': row.get('증빙유형', ''),
                    'note': row.get('적요', ''),
                    'project': row.get('프로젝트', ''),
                    'start_date': start_date,
                    'end_date': end_date,
                    'original_data': row.to_dict()
                }
                processed_data.append(processed_row)
            
            return processed_data
            
        except Exception as e:
            raise Exception(f"데이터 처리 중 오류: {e}")

    def clean_amount(self, amount):
        """금액에서 쉼표, 소숫점, 공백 제거하고 정수로 변환"""
        try:
            if isinstance(amount, (int, float)):
                # 숫자형이면 정수로 변환
                return str(int(amount))
            elif isinstance(amount, str):
                # 문자열이면 모든 특수문자 제거 후 정수 변환
                cleaned = amount.replace(',', '').replace(' ', '').replace('.0', '').replace('.00', '')
                # 소숫점이 있으면 소숫점 이하 제거
                if '.' in cleaned:
                    cleaned = cleaned.split('.')[0]
                # 빈 문자열이면 0 반환
                if not cleaned:
                    return "0"
                return str(int(float(cleaned)))
            else:
                return "0"
        except (ValueError, TypeError):
            return "0"
    
    def parse_date(self, date_str):
        """날짜 문자열 파싱"""
        try:
            # 다양한 날짜 형식 지원
            formats = ["%Y%m%d", "%Y-%m-%d", "%Y.%m.%d", "%m/%d/%Y"]
            for fmt in formats:
                try:
                    return datetime.strptime(str(date_str), fmt)
                except:
                    continue
            raise ValueError("지원하지 않는 날짜 형식")
        except:
            return datetime.now()