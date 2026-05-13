import pandas as pd
import openpyxl
from pathlib import Path

def split_column_by_delimiter(input_file, output_file=None):
    """
    엑셀 파일의 A열 데이터를 " - " (스페이스-하이픈-스페이스) 패턴으로만 분리하여 두 개의 시트로 저장합니다.
    
    Parameters:
    -----------
    input_file : str
        입력 엑셀 파일 경로
    output_file : str, optional
        출력 엑셀 파일 경로 (기본값: input_file에 '_processed' 접미사 추가)
    """
    
    if output_file is None:
        # 기본적으로 '_processed' 접미사를 붙여서 저장
        input_path = Path(input_file)
        base_output = input_path.parent / f"{input_path.stem}_processed{input_path.suffix}"
        
        # 파일이 이미 존재하면 번호를 붙여서 저장
        counter = 1
        output_file = base_output
        while output_file.exists():
            output_file = input_path.parent / f"{input_path.stem}_processed_{counter}{input_path.suffix}"
            counter += 1
    
    # 엑셀 파일 읽기 (첫 번째 시트)
    df_original = pd.read_excel(input_file, sheet_name=0)
    
    # A열 (첫 번째 열) 데이터 확인
    if df_original.empty:
        print("엑셀 파일이 비어있습니다.")
        return
    
    # A열을 기준으로 좌우 분리
    if len(df_original.columns) > 0:
        first_column = df_original.columns[0]
        
        # 분리된 데이터프레임 생성
        df_split = pd.DataFrame()
        
        # " - " (스페이스-하이픈-스페이스) 패턴으로만 분리하는 함수
        def split_by_spaced_hyphen(text):
            text = str(text).strip()
            # " - " 패턴이 있는지 확인
            if ' - ' in text:
                parts = text.split(' - ', 1)  # 최대 1번만 분리
                return parts[0].strip(), parts[1].strip() if len(parts) > 1 else None
            else:
                # 패턴이 없으면 전체를 A에 넣고 B는 None
                return text, None
        
        # 각 행에 대해 분리 적용
        split_results = df_original[first_column].apply(split_by_spaced_hyphen)
        df_split['A'] = split_results.apply(lambda x: x[0])
        df_split['B'] = split_results.apply(lambda x: x[1])
        
        # 각 열에서 중복 제거 및 빈 값 제거
        unique_a = df_split['A'].dropna().drop_duplicates().reset_index(drop=True)
        unique_b = df_split['B'].dropna().drop_duplicates().reset_index(drop=True)
        
        # 중복 제거된 데이터프레임 생성
        max_len = max(len(unique_a), len(unique_b))
        df_unique = pd.DataFrame({
            'A': pd.concat([unique_a, pd.Series([None] * (max_len - len(unique_a)))]).reset_index(drop=True),
            'B': pd.concat([unique_b, pd.Series([None] * (max_len - len(unique_b)))]).reset_index(drop=True)
        })
        
        # 결과 출력
        print("\n원본 데이터 (Sheet1):")
        print(df_original.head(5))
        print("\n분리된 데이터 (중복 제거 전):")
        print(df_split.head(5))
        print("\n중복 제거된 고유 데이터 (Sheet2):")
        print(df_unique.head(10))
        
        # ExcelWriter를 사용하여 여러 시트에 저장
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet1: 원본 데이터
            df_original.to_excel(writer, sheet_name='Sheet1', index=False)
            # Sheet2: 중복 제거된 고유 데이터
            df_unique.to_excel(writer, sheet_name='Sheet2', index=False)
        
        print(f"\n완료! 파일이 저장되었습니다: {output_file}")
        print("- Sheet1: 원본 데이터")
        print("- Sheet2: 중복 제거된 고유 데이터 (A열: 왼쪽 고유값, B열: 오른쪽 고유값)")
    else:
        print("열을 찾을 수 없습니다..")

# 사용 예시
if __name__ == "__main__":
    # 처리할 엑셀 파일 경로를 입력하세요
    # 현재 스크립트 파일의 디렉토리를 기준으로 경로 설정
    script_dir = Path(__file__).parent
    input_excel = script_dir / "Function_20260507.xlsx"  # ← 여기를 자신의 엑셀 파일로 변경
    
    split_column_by_delimiter(str(input_excel))
