import os
import shutil
from pathlib import Path

def copy_new_files(source_root, target_root):
    """
    D:\ 의 파일을 E:\01_Daimler_Work 로 복사
    기존 파일은 그대로 두고 새로운 파일만 복사
    
    Args:
        source_root: 복사 원본 폴더 (예: D:\)
        target_root: 복사 대상 폴더 (예: E:\01_Daimler_Work)
    """
    source_path = Path(source_root)
    target_path = Path(target_root)
    
    # 대상 폴더가 없으면 생성
    if not target_path.exists():
        target_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ 대상 폴더 생성: {target_path}")
    
    copied_count = 0
    skipped_count = 0
    
    # D:\ 의 모든 파일과 폴더 재귀적으로 탐색
    for source_item in source_path.rglob('*'):
        # 상대 경로 계산
        relative_path = source_item.relative_to(source_path)
        target_item = target_path / relative_path
        
        if source_item.is_file():
            # 파일인 경우
            if target_item.exists():
                # 대상 파일이 이미 존재하면 건너뛰기
                skipped_count += 1
                print(f"⊘ 건너뜸 (이미 존재): {relative_path}")
            else:
                # 대상 파일이 없으면 복사
                # 필요한 폴더 구조 생성
                target_item.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source_item, target_item)
                copied_count += 1
                print(f"✓ 복사됨: {relative_path}")
        
        elif source_item.is_dir():
            # 폴더인 경우
            if not target_item.exists():
                target_item.mkdir(parents=True, exist_ok=True)
                print(f"✓ 폴더 생성: {relative_path}")
    
    # 결과 출력
    print("\n" + "="*60)
    print(f"복사 완료!")
    print(f"  - 새로 복사된 파일: {copied_count}개")
    print(f"  - 건너뛴 파일 (기존): {skipped_count}개")
    print(f"  - 원본: {source_root}")
    print(f"  - 대상: {target_root}")
    print("="*60)


if __name__ == "__main__":
    source = r"D:\ ".strip()  # D:\ 의 모든 파일
    target = r"E:\_VW_Work"
    
    # 경로 존재 확인
    if not Path(source).exists():
        print(f"❌ 원본 경로가 존재하지 않습니다: {source}")
        exit(1)
    
    # 복사 실행
    try:
        copy_new_files(source, target)
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        exit(1)
