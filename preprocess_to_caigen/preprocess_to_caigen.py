"""
KMWE 연구용 골드 데이터 구축을 위한 전처리 스크립트
원본 문장을 형태소 단위로 분리하여 CAIGen(라벨링 툴)에 주입할 데이터 생성
"""

import os
import json
import pandas as pd
from kiwipiepy import Kiwi
from pathlib import Path


def ensure_data_directory():
    """data 폴더가 없으면 생성"""
    data_dir = Path("data")
    data_dir.mkdir(exist_ok=True)
    return data_dir


def create_dummy_excel_if_not_exists(excel_path):
    """raw_sentences.xlsx 파일이 없으면 더미 데이터로 생성"""
    if not os.path.exists(excel_path):
        print(f"{excel_path} 파일이 없습니다. 더미 데이터를 생성합니다...")
        dummy_data = {
            "target_sentence": [
                "학교에 갔다",
                "나는 어제 친구를 만났다",
                "오늘 날씨가 좋다",
                "책을 읽고 있다"
            ]
        }
        df = pd.DataFrame(dummy_data)
        df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"더미 데이터가 {excel_path}에 생성되었습니다.")


def load_sentences(excel_path):
    """엑셀 파일에서 문장 목록 로드"""
    df = pd.read_excel(excel_path, engine='openpyxl')
    
    # target_sentence 컬럼 확인
    if 'target_sentence' not in df.columns:
        raise ValueError(f"{excel_path}에 'target_sentence' 컬럼이 없습니다.")
    
    # NaN 값 제거
    sentences = df['target_sentence'].dropna().tolist()
    return sentences


def tokenize_sentences(sentences):
    """Kiwi 형태소 분석기를 사용하여 문장들을 토큰화"""
    # Kiwi 모델 초기화 (sbg 모델 사용 시도, 없으면 기본 모델 사용)
    try:
        kiwi = Kiwi(model_type='sbg')
        print("sbg 모델을 사용합니다.")
    except (ValueError, TypeError, Exception) as e:
        # sbg 모델이 없거나 지원되지 않으면 기본 모델 사용
        print(f"sbg 모델을 사용할 수 없습니다. 기본 모델을 사용합니다. (오류: {e})")
        kiwi = Kiwi()
    
    tokenized_results = []
    
    for sentence in sentences:
        # 문장을 문자열로 변환 (NaN 등 처리)
        sentence_str = str(sentence).strip()
        if not sentence_str or sentence_str == 'nan':
            continue
        
        # 형태소 분석 수행 (첫 번째 후보만 사용)
        # kiwipiepy의 analyze()는 (tokens, score) 튜플의 리스트를 반환
        analysis_results = kiwi.analyze(sentence_str)
        if not analysis_results or len(analysis_results) == 0:
            print(f"경고: '{sentence_str}'에 대한 분석 결과가 없습니다.")
            continue
        
        # 첫 번째 분석 후보의 토큰 리스트와 점수 추출
        tokens, score = analysis_results[0]
        
        # 각 Token 객체에서 형태소(form)와 char-offset(start, end) 추출
        # Kiwi의 Token 객체는 start와 end 속성을 제공하므로 이를 직접 사용
        # 이렇게 하면 공백을 포함한 정확한 offset 계산 가능
        token_list = []
        
        for token in tokens:
            token_form = token.form
            
            # Kiwi Token 객체의 start, end 속성 사용 (공백 포함 정확한 위치)
            # hasattr로 속성 존재 여부 확인 후 사용
            if hasattr(token, 'start') and hasattr(token, 'end'):
                token_start = token.start
                token_end = token.end
            else:
                # start/end 속성이 없는 경우 (구버전 호환성)
                # 원본 문장에서 형태소를 찾아서 offset 계산
                found_pos = sentence_str.find(token_form)
                if found_pos != -1:
                    token_start = found_pos
                    token_end = found_pos + len(token_form)
                else:
                    # 찾지 못한 경우 경고 출력
                    print(f"경고: '{token_form}'를 문장 '{sentence_str}'에서 찾을 수 없습니다.")
                    token_start = 0
                    token_end = len(token_form)
            
            token_obj = {
                'surface': token_form,
                'start': token_start,
                'end': token_end
            }
            
            token_list.append(token_obj)
        
        tokenized_results.append({
            'sentence': sentence_str,
            'tokens': token_list
        })
    
    return tokenized_results


def create_tokenized_excel(tokenized_results, output_path):
    """
    토큰화 결과를 엑셀 파일로 저장 (원문 + 형태소들이 각 셀에)
    - A열: target_sentence (원문)
    - B열 ~ N열: 분리된 형태소가 각 셀에 하나씩 순서대로
    """
    if not tokenized_results:
        print("경고: 토큰화된 결과가 없어 엑셀 파일을 생성할 수 없습니다.")
        return
    
    # 최대 형태소 개수 찾기 (컬럼 수 결정)
    max_tokens = max(len(result['tokens']) for result in tokenized_results)
    
    # 데이터 프레임 생성용 리스트
    # A열: 원문 문장
    data_dict = {'target_sentence': [result['sentence'] for result in tokenized_results]}
    
    # B열부터: 형태소들이 각 셀에 하나씩 순서대로
    for i in range(max_tokens):
        col_name = f'token_{i+1}'
        data_dict[col_name] = []
        for result in tokenized_results:
            tokens = result['tokens']
            if i < len(tokens):
                # token이 딕셔너리인 경우 surface 추출, 문자열인 경우 그대로 사용
                if isinstance(tokens[i], dict):
                    data_dict[col_name].append(tokens[i].get('surface', ''))
                else:
                    data_dict[col_name].append(tokens[i])
            else:
                data_dict[col_name].append('')
    
    df = pd.DataFrame(data_dict)
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"토큰화 결과가 {output_path}에 저장되었습니다.")


def create_caigen_json(tokenized_results, output_path):
    """
    CAIGen_KrCh용 JSON 파일 생성
    구조: List of Objects
    [
      {
        "id": "1",                    // 1부터 시작하는 순차 ID
        "target_sentence": "학교에 갔다", // 원문 (target_sentence 필드 사용)
        "tokens": [                   // 형태소 리스트
          {"surface": "학교", "start": 0, "end": 2},
          {"surface": "에", "start": 2, "end": 3},
          ...
        ]
      }
    ]
    """
    if not tokenized_results:
        print("경고: 토큰화된 결과가 없어 JSON 파일을 생성할 수 없습니다.")
        return
    
    json_data = []
    
    # 1부터 시작하는 순차 ID로 변환
    for idx, result in enumerate(tokenized_results, start=1):
        # tokens가 이미 딕셔너리 리스트인 경우 그대로 사용
        # (start, end가 포함된 형태)
        if result['tokens'] and isinstance(result['tokens'][0], dict):
            tokens = result['tokens']
        else:
            # 기존 형식 (문자열 리스트)인 경우 변환
            tokens = [{"surface": token} for token in result['tokens']]
        
        json_obj = {
            "id": str(idx),
            "target_sentence": result['sentence'],  # target_sentence 필드 사용
            "tokens": tokens
        }
        json_data.append(json_obj)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    
    print(f"CAIGen_KrCh용 JSON 파일이 {output_path}에 저장되었습니다.")


def main():
    """
    메인 실행 함수
    - data 폴더 확인 및 생성
    - raw_sentences.xlsx 파일 로드 (없으면 더미 데이터 생성)
    - 형태소 분석 수행
    - tokenized_view.xlsx 및 upload_data.json 생성
    """
    # data 폴더 확인 및 생성
    data_dir = ensure_data_directory()
    
    # 입력 파일 경로
    input_excel = data_dir / "raw_sentences.xlsx"
    
    # 더미 데이터 생성 (파일이 없는 경우)
    create_dummy_excel_if_not_exists(input_excel)
    
    # 문장 로드
    print("원본 문장을 로드하는 중...")
    try:
        sentences = load_sentences(input_excel)
        print(f"총 {len(sentences)}개의 문장을 로드했습니다.")
    except Exception as e:
        print(f"오류: 문장을 로드하는 중 문제가 발생했습니다: {e}")
        return
    
    # 형태소 분석
    print("형태소 분석을 수행하는 중...")
    tokenized_results = tokenize_sentences(sentences)
    print(f"형태소 분석이 완료되었습니다. (처리된 문장 수: {len(tokenized_results)})")
    
    # 출력 파일 경로
    output_excel = data_dir / "tokenized_view.xlsx"
    output_json = data_dir / "upload_data.json"
    
    # 토큰화 결과 엑셀 파일 생성
    create_tokenized_excel(tokenized_results, output_excel)
    
    # CAIGen용 JSON 파일 생성
    create_caigen_json(tokenized_results, output_json)
    
    print("\n전처리 작업이 완료되었습니다!")
    print(f"- 토큰화 결과 엑셀: {output_excel}")
    print(f"- CAIGen용 JSON: {output_json}")


if __name__ == "__main__":
    main()
