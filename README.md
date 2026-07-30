# Excel → Datasphere 업로드 마법사

[![tests](https://github.com/jhanseo/datasphere-excel-loader/actions/workflows/tests.yml/badge.svg)](https://github.com/jhanseo/datasphere-excel-loader/actions/workflows/tests.yml)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/downloads/)

엑셀 파일의 특정 시트를 SAP Datasphere(HANA) Open SQL Schema 테이블로 올려주는 도구입니다.
현업 담당자가 직접 쓸 수 있도록 GUI 마법사로 만들었고, 배치 자동화를 위한 CLI도 함께 제공합니다.

시트 이름이 정해진 키워드로 시작하는 시트만 골라 테이블로 올리며, 테이블이 없으면 데이터
유형을 자동으로 추론해 만들어 줍니다. 업로드 전에 **미리보기**로 컬럼별 데이터 유형과
설명(COMMENT)을 확인하고 직접 바꿀 수 있고, **점검**으로 바꾼 유형이 실제 데이터와 맞는지
검사할 수 있습니다. 사내 DRM(문서보안)이 걸린 엑셀 파일도 설치된 Excel을 경유해 읽습니다.

> **참고**  
> 이 도구는 사내 SAP Datasphere 환경을 전제로 만들어졌습니다. 사용하려면 Open SQL Schema
> 권한이 있는 데이터베이스 사용자가 필요합니다. DRM 파일 처리는 Microsoft Excel이 설치된
> Windows에서만 동작합니다. 그 외 기능은 macOS·Linux에서도 그대로 쓸 수 있습니다.

## 주요 기능

- 엑셀 파일 여러 개를 한 번에 선택해 업로드 (드래그 앤 드롭 지원)
- 헤더 행 자동 감지, 헤더 다음 행을 컬럼 COMMENT로 등록
- 데이터 유형 자동 추론 + 미리보기 화면에서 직접 수정
- 컬럼별 `PRIMARY KEY` / `NOT NULL` 지정
- 적재 모드 3종: `TRUNCATE` / `INSERT` / `UPSERT`
- 유형 불일치 시 처리 방식 선택 (NULL 변환 · 행 제외 · 유형 되돌리기 · 중단)
- DRM(DRMONE) 파일 감지 및 Excel 경유 읽기
- 업로드 실패 시 기존 데이터 보존 (TRUNCATE 롤백)
- `config.ini`의 모든 항목을 GUI 탭에서 편집
- 단일 실행 파일(exe) 빌드 스크립트 포함

## 구성 파일

| 파일 | 설명 |
| --- | --- |
| `wizard.py` | PyQt5 기반 GUI 마법사 (실행 진입점) |
| `excel_loader.py` | 엑셀 읽기 · 타입 추론 · 검증 · HANA 업로드 엔진 (CLI 겸용) |
| `config.ini.example` | 설정 파일 예시. `config.ini`로 복사해 사용합니다 |
| `build_exe.bat` | (Windows용) 의존 패키지 설치부터 exe 빌드까지 한 번에 수행 |
| `wizard.spec` | PyInstaller 빌드 스펙 |
| `requirements.txt` | 필요 패키지 목록 |
| `tests/` | pytest 테스트 (실제 DB·Excel·PyQt5 없이 동작) |

## 설치와 실행

```
git clone https://github.com/jhanseo/datasphere-excel-loader.git
cd datasphere-excel-loader
pip install -r requirements.txt
cp config.ini.example config.ini      # Windows: copy config.ini.example config.ini
python wizard.py
```

`config.ini`에는 데이터베이스 비밀번호가 평문으로 저장되므로 `.gitignore`에 등록되어
있습니다. 저장소에 커밋되지 않도록 주의하세요.

### 배포용 실행 파일 생성
Windows에서 `build_exe.bat`을 실행합니다.
빌드가 끝나면 `dist\ExcelUploadWizard\` 폴더가 만들어지고 그 안에
`ExcelUploadWizard.exe`와 `config.ini`가 들어갑니다. **폴더째 zip으로 묶어** 배포하면
되고, 사용자는 압축을 푼 뒤 exe를 실행하면 됩니다. 사용자 PC에는 Python이 필요 없지만,
DRM이 걸린 파일을 다루려면 Microsoft Excel이 설치되어 있어야 합니다.

`wizard.spec` 맨 위의 `ONEFILE = False`를 `True`로 바꾸면
`dist\ExcelUploadWizard.exe` 하나만 생성됩니다.
해당 옵션의 경우, 실시간 검사로 인해 빌드가 실패할 수 있습니다.
DRM 사용 환경일 경우, 백신 오탐을 줄이고 실행 속도를 높이기 위해 기본 값인 폴더 방식을 권합니다.

### build_exe.bat이 아예 실행되지 않을 때

명령 프롬프트에서 아래 세 줄을 직접 실행합니다.

```
cd /d "이 폴더의 경로"
python -m pip install -r requirements.txt
python -m PyInstaller --clean --noconfirm wizard.spec
```

## 화면 구성

마법사는 세 단계의 작업 탭과 다섯 개의 설정 탭으로 나뉩니다.

**1. 파일 선택** 탭에서 업로드할 엑셀 파일을 여러 개 고릅니다. 파일 선택 대화상자로
추가하거나 창에 끌어다 놓을 수 있습니다. DRM(문서보안)이 걸린 파일은 목록에서 🔒 표시와
함께 주황색으로 나타납니다.

**설정 탭**은 `config.ini`의 섹션별로 나뉘어 있습니다. HANA 접속(호스트·포트·사용자·
비밀번호·스키마·프록시), 일반(시트 키워드, 적재 모드, 파일명 prefix, Excel 창 표시,
기본 폴더), 헤더(자동 감지, 헤더 행 번호, 설명 행 사용, 미리보기 행 수), 로깅(폴더, 레벨),
보관(처리 완료 파일 이동, archive 폴더)까지 모든 항목을 GUI에서 편집할 수 있습니다.
하단의 **[설정 저장]** 버튼을 누르면 `config.ini`에 반영됩니다.

**2. 미리보기** 탭에서는 왼쪽 트리에 파일과 대상 시트가, 오른쪽에 해당 시트의 컬럼 목록과
데이터 샘플이 표시됩니다. 컬럼 목록에는 컬럼명, 원본 컬럼명, 자동 추론된 데이터 유형,
지정 유형, KEY, NOT NULL, Comment, NULL 개수, 샘플 값이 나옵니다. 지정 유형과 Comment는
직접 수정할 수 있고, '사용' 체크를 해제하면 그 컬럼은 업로드에서 제외됩니다. 트리의 시트
체크박스를 끄면 시트 단위로 제외됩니다.

**[KEY]** 를 체크한 컬럼들은 `PRIMARY KEY`가 됩니다. 여러 컬럼을 체크하면 체크한 순서대로
복합 키가 만들어집니다. HANA에서 PRIMARY KEY는 자동으로 `NOT NULL`이므로, KEY를 체크하면
NOT NULL도 함께 켜지고 해제할 수 없습니다. **[NOT NULL]** 은 단독으로도 지정할 수 있습니다.
두 제약은 **테이블을 새로 만들 때(CREATE TABLE)만 반영됩니다.**

**3. 실행 로그** 탭에는 진행 상황과 결과가 실시간으로 표시되며, 같은 내용이 로그 폴더에도
일자별 파일로 쌓입니다.

## 버튼 동작

**[미리보기]** 는 설치된 Excel을 통해 파일을 여는 DRM 우회 경로를 사용합니다. DRM 파일이
포함되어 있으면 Excel이 실행되고, 보안 로그인이나 열기 승인 창이 뜨면 사용자가 완료해 주면
됩니다. 이렇게 읽어온 데이터는 메모리에 캐시되어 이어지는 [실행]에 그대로 사용되므로,
DRM 파일도 별도의 복호화 없이 업로드할 수 있습니다.

**[점검]** 은 미리보기에서 데이터 유형을 바꾼 컬럼을 대상으로 실제 변환이 가능한지 검사합니다.
변환할 수 없는 값이 발견되면 어느 시트의 어느 컬럼에서 몇 건이 문제인지, 예시 값은 무엇인지
보여주는 대화상자가 열리고, 컬럼별로 처리 방식을 고르게 합니다. 선택지는 네 가지입니다.
`NULL로 변환 후 진행`은 실패한 셀만 빈 값으로 넣고 나머지는 정상 적재합니다.
`해당 행 제외`는 문제가 있는 셀이 포함된 행 전체를 업로드에서 뺍니다.
`원래(추론) 유형으로 되돌리기`는 자동 추론된 안전한 유형으로 롤백해 데이터 손실 없이 적재합니다.
`업로드 중단`은 불일치가 하나라도 있으면 그 시트를 오류로 처리합니다.
대화상자 상단의 일괄 적용으로 모든 컬럼에 같은 방식을 한 번에 지정할 수도 있습니다.

[KEY] 또는 [NOT NULL]을 체크한 경우 `[점검]`이 추가로 두 가지를 확인합니다. 먼저 데이터가
제약을 만족하는지 봅니다. NOT NULL 컬럼에 빈 값이 있거나 KEY 컬럼에 빈 값 또는 중복이
있으면 위반 목록과 문제가 된 행 번호를 보여주고 점검을 통과시키지 않습니다. 그대로
올리면 실패할 것이기 때문에, 체크를 해제하거나 엑셀을 고친 뒤 다시 점검해야 합니다.
다음으로 대상 테이블이 이미 있는지 조회합니다. 이미 있다면 제약을 적용할 수 없으므로
지정한 KEY/NOT NULL과 테이블의 실제 PRIMARY KEY를 나란히 보여주며 경고합니다.
UPSERT 모드인데 PRIMARY KEY가 없는 경우도 이 단계에서 걸러집니다. HANA 접속 정보와
스키마가 비어 있으면 이 조회는 생략되고 안내만 표시됩니다.

**[실행]** 은 미리보기 캐시가 있으면 그 데이터를 그대로 업로드합니다. 캐시가 없는 상태에서
DRM이 걸린 파일이 하나라도 포함되어 있으면, 미리보기 사용을 안내하는 문구를 띄우고
**전체 실행을 중단**합니다. DRM이 없는 파일만 선택한 경우에는 미리보기 없이 바로 실행해도
정상적으로 업로드됩니다. 유형을 바꿨는데 점검을 하지 않은 컬럼이 있으면 실행 직전에 점검
여부를 한 번 더 확인합니다.

**[모의 실행 (Dry-run)]** 은 DB에 접속하지 않고 어떤 시트가 어떤 테이블로, 몇 건이 올라갈지만
로그로 보여줍니다. 설정을 처음 맞출 때 사용하면 좋습니다.

## 적재 모드

`[⚙ 일반]` 탭의 **적재 모드**로 세 가지 중 하나를 고릅니다.

| 모드 | 동작 | 실행되는 SQL |
| --- | --- | --- |
| `truncate` | 기존 데이터를 모두 지우고 새로 적재 (기본) | `TRUNCATE` → `INSERT` |
| `insert` | 기존 데이터를 남기고 뒤에 추가 | `INSERT` |
| `upsert` | 키가 같으면 갱신, 없으면 추가 | `UPSERT ... WITH PRIMARY KEY` |

`insert`는 같은 파일을 두 번 올리면 중복 적재되므로 주의해야 합니다. `upsert`는 대상
테이블에 `PRIMARY KEY`가 반드시 있어야 하며, 없으면 실행이 실패합니다. 테이블을 새로
만드는 경우에는 미리보기에서 **[KEY]** 로 체크한 컬럼이 `PRIMARY KEY`가 되고, 이미
있는 테이블이라면 그 테이블에 정의된 키가 사용됩니다. `[점검]`이 이 조건을 미리
확인해 줍니다.

## 업로드 안전장치

**기존 데이터가 유실되지 않습니다.** `truncate` 모드에서 HANA의 `TRUNCATE`는 DML이 아니라
DDL입니다. 세션의 `AUTOCOMMIT DDL`이 켜져 있으면 `TRUNCATE`가 즉시 커밋되어, 그 뒤
INSERT가 실패하고 rollback을 해도 기존 데이터는 이미 사라진 상태가 됩니다.
`conn.setautocommit(False)`는 DML에만 적용되므로 이것만으로는 막을 수 없습니다.

그래서 접속 직후 `SET TRANSACTION AUTOCOMMIT DDL OFF`를 실행해 `TRUNCATE`까지 롤백
대상에 넣습니다. 이 덕분에 업로드가 중간에 실패해도 기존 데이터는 그대로 남고, 로그에
그 사실이 기록됩니다. 이 설정이 권한 문제로 실패하면 적재는 그대로 진행하되, 복구가
불가능하다는 경고를 로그에 남깁니다. 실제 실행된 문장은 로그의 `[DML]` 줄에서 확인할
수 있습니다.

**이미 존재하는 테이블은 유형·KEY·NOT NULL 변경이 적용되지 않습니다.** 기존 테이블에는
`CREATE`를 다시 하지 않으므로, 미리보기에서 바꾼 데이터 유형과 체크한 [KEY]/[NOT NULL]은
반영되지 않고 테이블에 이미 정의된 것이 그대로 쓰입니다. `[점검]`과 업로드 직전 확인
단계에서 대상 테이블의 실제 컬럼 유형과 PRIMARY KEY를 조회해, 지정한 것과 다르면 비교
표를 띄워 계속할지 취소할지 묻습니다. 반드시 적용해야 한다면 취소하고 Datasphere에서
컬럼 유형이나 제약을 변경하거나 테이블을 삭제한 뒤 다시 실행하세요.

## 미리보기 색상 표시

컬럼 행의 배경색은 상태를 나타냅니다. 흰색은 자동 추론 유형 그대로인 컬럼,
노란색은 유형을 바꿨지만 아직 점검하지 않은 컬럼, 초록색은 점검을 통과한 컬럼,
빨간색은 점검에서 불일치가 발견되어 처리 방식이 지정된 컬럼입니다.

## CLI 사용법

GUI 없이 배치로 돌릴 때는 기존과 동일하게 `excel_loader.py`를 사용합니다.

```
python excel_loader.py --config config.ini
python excel_loader.py --config config.ini --dry-run
python excel_loader.py --config config.ini --file A.xlsx --file B.xlsx
python excel_loader.py --config config.ini --only-table T_SALES,T_CUSTOMER
python excel_loader.py --config config.ini --allow-drm
python excel_loader.py --config config.ini --mode insert
```

`--mode`로 config의 `load_mode`를 덮어쓸 수 있습니다 (`truncate` / `insert` / `upsert`).
CLI에는 미리보기 화면이 없으므로 [KEY]/[NOT NULL] 지정은 GUI에서만 가능합니다.

`--file`을 주지 않으면 `[GENERAL] excel_dir` 폴더의 모든 엑셀 파일을 처리합니다.
CLI는 기본적으로 DRM 파일을 거부하며, `--allow-drm`을 붙이면 GUI의 미리보기와 같은
Excel 우회 경로를 사용합니다.

## 자주 겪는 문제

DRM 파일 미리보기가 실패한다면 대상 PC에 Excel이 설치되어 있는지, `xlwings`와 `pywin32`가
설치되어 있는지 확인하세요. HANA 접속이 되지 않으면 사내 방화벽에서 해당 호스트의 443
아웃바운드가 열려 있는지, 프록시가 필요한 환경인지, Datasphere IP 허용 목록에 현재 공인 IP가
등록되어 있는지를 순서대로 확인하면 됩니다. 대상 시트를 찾지 못한다는 안내가 나오면
`[⚙ 일반]` 탭의 시트 이름 키워드가 실제 시트명 접두어와 맞는지 확인하세요.

## 개발

테스트는 실제 HANA·Excel·PyQt5 없이 돌아갑니다. HANA 접속은 가짜 커넥션으로,
GUI는 `tests/mockqt`의 Qt 대역으로 대체하고, 샘플 엑셀은 테스트가 직접 만듭니다.

```
pip install pandas openpyxl pytest
python -m pytest tests/ -v
```

커버하는 범위는 데이터 유형 추론과 표기 검증, 불일치 처리 정책 4종, KEY/NOT NULL 제약
검사, 적재 모드별 SQL과 DDL 생성, DRM 감지와 실행 차단, 그리고 GUI의 시그널·대화상자
흐름입니다. 실제 Qt 렌더링과 xlwings의 DRM 우회는 Windows에서 수동으로 확인해야 합니다.

기여를 환영합니다. 이슈나 Pull Request를 남겨 주세요. 코드를 수정할 때는 `pytest`가
통과하는지, `pyflakes`에 걸리는 것이 없는지 확인해 주시면 좋습니다.

## 라이선스

[MIT](LICENSE)

## 참고
- *HANA Database 사용자 생성* : 
    https://help.sap.com/docs/SAP_DATASPHERE/be5967d099974c69b77f4549425ca4c0/798e3fd6707940c3bd2219b2d1ebaac2.html?locale=en-US#create-a-database-user-with-password-based-authentication

- *Python 설치* : 

    (Windows) https://www.python.org/downloads/windows/
    (MacOS) https://www.python.org/downloads/macos/
