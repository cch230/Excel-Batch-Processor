# 엑셀 일괄배송처리 프로그램  
# Excel Batch Processor

---

## 🙏 후원 안내 (Support & Sponsor)

이 프로젝트가 도움이 되었다면, 개발 지속과 유지보수를 위해 후원을 부탁드립니다!  
여러분의 작은 응원이 오픈소스 발전에 큰 힘이 됩니다.

- [GitHub Sponsors로 후원하기](https://github.com/sponsors/cch230)
- 또는 커피 한 잔을 보내주세요! ☕

If you find this project useful, please consider supporting it!  
Your sponsorship helps keep this project alive and motivates further development.

- [Sponsor via GitHub Sponsors](https://github.com/sponsors/cch230)
- Or just buy me a coffee! ☕

감사합니다! Thank you!

---

### 1. 소개
[스마트스토어 / 쿠팡 / 토스] 주문 데이터와 [로젠택배 / CJ대한통운] 운송장 데이터를 자동 매칭하여 엑셀로 저장하는 PyQt5 기반 GUI 프로그램입니다. 암호화된 엑셀도 지원하며, 드래그 앤 드롭으로 파일을 올릴 수 있습니다.  
[Click to read English introduction.](#English-Guide)

   
### 2. 주요 기능
- **드래그 앤 드롭**으로 주문/운송장 엑셀 업로드
- **암호화 엑셀 지원** (`msoffcrypto` 사용)
- **수취인명, 전화번호, 주소 기준 자동 매칭**
- **결과 테이블 미리보기 및 엑셀 저장**
- **헤더 스타일(폰트, 배경색 등) 자동 적용**

### 3. 설치 방법
##### 필수 조건
- Python 3.7 이상
- pip

##### 의존성 설치
  ```bash
  pip install PyQt5 qasync pandas openpyxl msoffcrypto  
 ```

### 4. 사용 방법
1. 프로그램 실행
python delivery_ui.py
2. 주문/운송장 엑셀 파일을 각각 드래그 앤 드롭(또는 클릭)하여 선택
3. "일괄처리 시작" 클릭
4. 매칭 결과가 테이블에 표시되고, 엑셀로 저장됨 (`일괄처리_[A파일명].xlsx`)
5. 저장된 파일을 열지 여부 안내

### 5. 파일 포맷
#### 주문 데이터(A)
- 상품주문번호, 배송방법, 택배사, 상품명, 수량, 수취인명, 수취인연락처1, 통합배송지

#### 운송장 데이터(B)
- 수하인명, 수하인전화, 수하인주소1, 운송장번호

### 6. 기술 정보
- **GUI**: PyQt5
- **비동기**: qasync
- **엑셀 암호 해독**: msoffcrypto
- **엑셀 스타일**: openpyxl (헤더 굵게, 흰색 폰트, 파란 배경, 가운데 정렬)

### 7. 라이선스
GPL-3.0  
자세한 내용은 `LICENSE` 파일 참고


### 8. 예시 폴더 구조
  ```bash
/logen_batch_delivery  
├── delivery_ui.py  
├── requirements.txt   
├── LICENSE
├── icon.ico  
└── README.md  
 ```

## 9. 버전
- v1.0.0: 최초 배포
- v1.1.0: 암호화 엑셀 지원 추가
- v1.2.0: UI 개선 및 성능 최적화
- v1.3.5: 스마트 스토어, 쿠팡, 토스 - 로젠택배, CJ대한통운 
---

## English Guide

### 1. Introduction
A GUI tool for batch matching [Smart Store / Toss / Coupang] order and [LOZEN / CJ] shipping Excel files. Built with PyQt5, supports drag & drop, password-protected Excel, and saves results with styled headers.

### 2. Features
- **Drag & Drop** upload for order/shipping Excel files
- **Encrypted Excel support** (`msoffcrypto`)
- **Auto-matching** by name, phone, and address
- **Result table preview and Excel export**
- **Styled headers** (font, background color, alignment)

### 3. Installation
#### Requirements
- Python 3.7+
- pip

#### Install dependencies
  ```bash
  pip install PyQt5 qasync pandas openpyxl msoffcrypto  
 ```

### 4.Usage
1. Run the program:
2. Drag and drop (or click) to select order and shipping Excel files
3. Click "Start Batch Processing"
4. Matching results are shown in the table and saved as Excel (`일괄처리_[A_filename].xlsx`)
5. Prompt to open the saved file

### 5. File Formats
##### Order Data (A)
- 상품주문번호 (Order No), 배송방법 (Delivery), 택배사 (Courier), 상품명 (Product), 수량 (Qty), 수취인명 (Recipient), 수취인연락처1 (Phone), 통합배송지 (Address)

##### Shipping Data (B)
- 수하인명 (Recipient Name), 수하인전화 (Phone), 수하인주소1 (Address), 운송장번호 (Tracking No)

### 6. Technical Details
- **GUI**: PyQt5
- **Async**: qasync
- **Excel Decryption**: msoffcrypto
- **Excel Styling**: openpyxl (bold header, white font, blue background, centered)

### 7. License
GPL-3.0  
See `LICENSE` for details

### 8. Example Structure
 ```bash
/SmartStore-Excel-Processor  
├── delivery_ui.py  
├── requirements.txt   
├── LICENSE
├── icon.ico  
└── README.md  
 ```

### 9. version
- v1.0.0: Initial release
- v1.1.0: Added encrypted Excel file support
- v1.2.0: UI improvements and performance optimization
- v1.3.5: Added [Smart Store / Toss / Coupang] order and [LOZEN / CJ] shipping Excel file support

