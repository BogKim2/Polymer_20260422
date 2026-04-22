---
name: quotation-generator
description: "견적서 자동 생성. '견적서 만들어줘', '견적 뽑아줘', '단가표 작성', '비용 산출', '견적서 작성' 요청 시 자동 트리거."
---

## 목적

수신 회사명과 품목 목록을 입력받아 케이에스모듈테크 견적서를 생성합니다.
Excel(.xlsx)과 DOCX(.docx) 두 파일을 동시에 생성합니다.

## 공급자 고정 정보

| 항목 | 값 |
|------|-----|
| 상호명 | (주) 케이에스모듈테크 |
| 사업자번호 | 808-81-03318 |
| 대표 | 김 복 기 |
| 주소 | 부산시 금정구 부산대학로 63번길 2, 부산대학교 물리학과(장전동) 307호 |
| 업태 | 제조업, 판매업, 설계업, 통신판매업 |
| 직인 | `E:/venture/proposal/JBLab/2nd/직인.png` |

## 절차

### 1단계: 정보 수집

다음 정보를 사용자에게 확인합니다.

- **수신처**: 회사명 (예: "(주) 홍길동산업")
- **품목 목록**: 각 품목의 `품목명 / 단위 / 수량 / 단가`
- **견적 번호** (선택, 없으면 오늘 날짜로 자동 생성)
- **출력 파일명** (선택, 없으면 `견적서_수신처명` 자동 사용)

### 2단계: Python 스크립트 수정 후 실행

`E:/venture/proposal/JBLab/2nd/1quotation/make_xlsx.py` 를 수정하여 실제 데이터를 삽입합니다.

**수정 위치:**
- `row 7` 수신처명: `merge(7, 3, 7, 16, val="수신처명", ...)` 부분
- 품목 행 (rows 13-40): `ws.cell(row=r, column=1).value` 등으로 품목 데이터 삽입
- 날짜: `row 5` 년/월/일 셀에 오늘 날짜 삽입
- 출력 경로: `out = "..."` 변수

**품목 데이터 삽입 패턴 (스크립트 하단에 추가):**

```python
# 품목 데이터 삽입
items = [
    ("품목명1", "EA", 1, 100000),
    ("품목명2", "SET", 2, 250000),
]
for i, (name, unit, qty, price) in enumerate(items):
    r = 13 + i
    ws.cell(row=r, column=1).value = name    # 품목명 (A열)
    ws.cell(row=r, column=14).value = unit   # 단위 (N열)
    ws.cell(row=r, column=18).value = qty    # 수량 (R열)
    ws.cell(row=r, column=21).value = price  # 단가 (U열)

# 날짜
import datetime
today = datetime.date.today()
ws.cell(row=5, column=3).value = today.year
ws.cell(row=5, column=7).value = today.month
ws.cell(row=5, column=11).value = today.day

# 수신처명
ws.cell(row=7, column=3).value = "수신처명"
```

실행:
```bash
cd "E:/venture/proposal/JBLab/2nd/1quotation"
python make_xlsx.py
```

### 3단계: DOCX 스크립트 수정 후 실행

`E:/venture/proposal/JBLab/2nd/1quotation/make_docx.js` 를 수정합니다.

**수정 위치:**
- `recipientRow` 의 TextRun text에 수신처명 삽입
- `itemRows` 를 실제 품목 데이터로 교체
- `sumRow` 합계금액 셀에 실제 합계 삽입
- `totalRow` 합계행 한글 금액 표기 삽입

실행:
```bash
cd "E:/venture/proposal/JBLab/2nd/1quotation"
node make_docx.js
```

### 4단계: 직인 이미지 삽입 (xlsx)

openpyxl 로 직인.png를 공급자 정보 영역 우측에 삽입합니다.

```python
from openpyxl.drawing.image import Image as XLImage
img = XLImage("E:/venture/proposal/JBLab/2nd/직인.png")
img.width = 80
img.height = 80
ws.add_image(img, "AN6")  # 대표 행 옆 (row 6, AN 컬럼)
wb.save(out)
```

### 5단계: 금액 검증

- 공급가액 합계 = 각 품목의 (수량 × 단가) 합산
- 세액 = 공급가액 × 10% (ROUND 처리)
- 총액 = 공급가액 + 세액

## 출력 파일

| 파일 | 경로 |
|------|------|
| Excel | `E:/venture/proposal/JBLab/2nd/1quotation/견적서_{수신처명}.xlsx` |
| Word  | `E:/venture/proposal/JBLab/2nd/1quotation/견적서_{수신처명}.docx` |

## 자체 검증 체크리스트

- [ ] 모든 품목 (수량 × 단가) 합계가 정확한가?
- [ ] 세액(VAT 10%)이 올바르게 계산되었는가?
- [ ] 수신처 회사명이 정확히 입력되었는가?
- [ ] 오늘 날짜가 삽입되었는가?
- [ ] 직인.png 위치가 대표자명 옆 올바른 위치인가?
- [ ] xlsx와 docx 두 파일 모두 생성되었는가?
- [ ] 파일명에 수신처명이 포함되었는가?
