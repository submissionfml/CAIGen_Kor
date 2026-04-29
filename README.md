# CAIGen_Kor 사용법  

## CAIGen_Kor 소개  

이 깃허브 페이지는 투고 논문 <언어 모델과 규칙을 활용한 한국어 문법 항목 자동 주석 연구>에서 스팬 자동 산출 목적으로 사용한 CAIGen_Kor의 사용법을 소개합니다. 이 깃허브 페이지는 현재 익명이지만 심사과정이 끝나면 심사결과에 상관없이 논문의 각주 표시와 함께 투고자의 깃허브 레포로 전체 공개로 전환될 예정입니다.  
CAIGen_Kor은 Ide, Y. 외(2025:27007)의 CAIGen 주석 인터페이스를 한국어에 맞게 수정하였습니다.   
기존 CAIGen에서는 스팬 자동 기록 기능은 없었지만 CAIGen_Kor는 한국어 형태소 분석기 Kiwi를 접목하여 문장에서 문법 항목을 선택하면 자동으로 스팬이 기록되게 업데이트하였습니다.  

CAIGen_Kor의 장점은 아래와 같습니다.
1. 불연속 문법 항목의 스팬을 쉽게 산출할 수 있습니다.
2. 한 문장에 문법 항목이 여러 번 나타날 때나 하나의 문법 항목 구성 요소 사이에 다른 문법 항목이 나타나도 쉽게 모든 문법 항목의 스팬을 산출할 수 있습니다. 
3. 구글드라이브를 사용하므로 연구자 간의 협업이 편리합니다.

아래 그림에서 CAIGen_Kor을 사용하여 실제로 스팬을 산출하는 방법을 볼 수 있습니다.   

![alt text](/image/CAIGen_Kor.png)


## 폴더 구성

```text
preprocess_to_caigen/
  preprocess_to_caigen.py
  requirements.txt
  data/
    raw_sentences.xlsx
    tokenized_view.xlsx
    upload_data.json
CAIGen_Kor.gs
README.md
```

## 실행 방법

1. 노트북이나 데스크탑 컴퓨터에서 형태소 분석 실행  

우선 이 깃허브 페이지의 상단에 초록색 `< >Code 버튼`을 클릭한 후 `Download ZIP`을 클릭하여 CAIGen_Kor의 필요 폴더와 파일을 모두 다운 받습니다  

preprocess_to_caigen/data/raw_sentences.xlsx 파일을 열고 target_sentence 열에 문법 항목 주석 대상 문장을 넣습니다.  

그 다음 파워쉘이나 터미널을 열고 명령어를 실행하기 위해 필요한 모듈을 설치합니다.   
아래 명령어를 실행합니다. "(preprocess_to_caigen 폴더가 있는 위치)"에는 "preprocess_to_caigen"폴더가 있는 폴더의 위치를 기록합니다.  

```text
윈도우os 명령어:

cd "(preprocess_to_caigen 폴더가 있는 위치)\preprocess_to_caigen"
py -m pip install -r requirements.txt

맥os 명령어:

cd "(preprocess_to_caigen 폴더가 있는 위치)/preprocess_to_caigen"
pip install -r requirements.txt
```

그 후에 문법 항목 주석 대상 문장들을 형태소 분석합니다. 
아래 명령어를 실행합니다.

```text
윈도우os 명령어:

cd "(preprocess_to_caigen 폴더가 있는 위치)\preprocess_to_caigen"
py preprocess_to_caigen.py

맥os 명령어:

cd "(preprocess_to_caigen 폴더가 있는 위치)/preprocess_to_caigen"
python3 preprocess_to_caigen.py
```

실행이 끝나면 아래 파일이 생성됩니다.

```text
preprocess_to_caigen/data/tokenized_view.xlsx
preprocess_to_caigen/data/upload_data.json
```

tokenized_view.xlsx는 연구자가 형태소 분석 결과를 눈으로 확인하기 위한 파일입니다.  
upload_data.json이 중요한 파일입니다. 형태소 분석 결과를 구글드라이브에 올려서 다음 과정을 수행할 파일입니다.


2. 구글드라이브 준비    
구글드라이브에 CAIGen_Kor를 실행하고 싶은 폴더를 하나 만들고 그 안에 아래 두 폴더를 만듭니다.

```text
json_files  
gs_files
```

그 다음 1번에서 만들었던 preprocess_to_caigen/data/upload_data.json 파일을 방금 만들었던 json_files 폴더에 업로드합니다.  


3. Google Apps Script 설치  

[https://script.google.com](https://script.google.com) 에서 새 프로젝트를 만듭니다.  

기본으로 생성되는 Code.gs 파일의 내용을 지우고, 저희 깃 내에 보이는 CAIGen_Kor.gs 파일 내용을 그대로 붙여넣습니다.  

그후, 아래 내용에 연구자의 json_files, gs_files 폴더가 있는 이 프로그램 루트 폴더 ID를 입력해야 합니다.  
구글드라이브 루트 폴더 인터넷 주소가 "https://drive.google.com/drive/u/0/folders/1l4ANmPeksTLNW4UMdnvleDzXK19FrfYP" 면 이 주소의 가장 마지막 부분인 "1l4ANmPeksTLNW4UMdnvleDzXK19FrfYP" 가 루트 폴더 ID입니다. 

```text
/********************
 아래에 json_files, gs_files 폴더가 있는 이 프로그램 루트 폴더 ID를 입력해야 합니다.
 구글드라이브 루트 폴더 인터넷 주소가 "https://drive.google.com/drive/u/0/folders/1l4ANmPeksTLNW4UMdnvleDzXK19FrfYP"면 이 주소의 가장 마지막 부부인 "1l4ANmPeksTLNW4UMdnvleDzXK19FrfYP"가 루트 폴더 ID입니다. 
 ********************/

var projectFolderId = "1l4ANmP00pTLNW4UMeajK5DzXK19FrfYP";
```


4. Google Apps Script 실행  
Apps Script 편집기에서 아래 순서대로 실행합니다.  
Apps Script 상단의 "실행" 버튼 왼쪽에 "드라이브에 프로젝트 저장"을 클릭합니다.  
Apps Script 상단의 "디버그" 버튼 오른쪽에 "createAndWriteSheets"가 보이면 그대로 두고 안 보이면 아래 화살표 표시를 눌러 선택합니다.  
그 후에 "실행"버튼을 클릭합니다.  
첫 실행시에 "승인 필요" 창이 나타날 수 있으며 "권한 검토"-"고급"-"제목 없는 프로젝트(으)로 이동(안전하지 않음)"을 차례대로 클릭하면 됩니다.  



5. 결과 확인  
실행이 끝나면 구글드라이브 gs_files 폴더 안에 스프레드시트가 생성됩니다.  
파일 이름은 `annotation_workbook`로 생성됩니다.

`annotation_workbook`스프레드시트를 클릭하시면 안에는 아래 시트가 생성됩니다.  

```text
facesheet  
upload_data  
CharOffset  
```

주석 스팬 작업은 `upload_data` 시트에서 하시면 됩니다.
주석 대상 형태소 사각형을 클릭하시면 형태소 위치에 대응되는 스팬이 좌측 상단의 Span, CharOffset 칸에 나타납니다.
스팬 작업의 총 결과가 `CharOffset`에서 나타납니다.

주의  
주석해야 하는 문장이 150개를 넘어가는 경우에는 `createAndWriteSheets`를 실행한 직후에 sample_annotator 파일의 시트 하단에서 문장이 계속 생성될 수 있습니다. 시트 생성이 충분히 끝난 뒤 주석을 시작하는 것을 권장합니다.  
문장 수가 많으면 한 번에 끝나지 않을 수 있습니다. 이 경우 writeSheetsResume()를 실행하면 이어서 작성합니다.  
Google Sheets 셀 길이 제한을 넘는 매우 긴 문장은 자동으로 건너뛰고 표시만 남길 수 있습니다.

공개본 점검 환경:
- Python 3.9.6
- kiwipiepy 0.22.2

## 라이선스  

CAIGen_Kor은 CAIGen(https://github.com/Yusuke196/CAIGen)을 기반으로 수정한 코드가 포함되어 있으며,  
해당 부분은 원저작물의 MIT License를 따릅니다. 이 저장소의 나머지 독자적 자료 및 추가 작성 부분은 비상업적 연구 및 교육 목적에 한하여 사용할 수 있습니다.

CAIGen_Kor은 한국어 형태소 분석을 위해 `kiwipiepy`를 사용합니다.  
`kiwipiepy`는 GNU Lesser General Public License v3 (LGPLv3)로 배포되므로, 사용자는 해당 라이선스 조건을 함께 준수해야 합니다.