# 장비데이터 CSV to EXCEL

## 삭제하면 안되는 파일

```
 ㄴ 양식.xlsx  // 엑셀 양식 파일
 ㄴ 01_input    // 입력 데이터 폴더 (작업 완료후 자동으로 삭제 후 생성됨)
 ㄴ 02_output   // 출력 데이터 폴더 (작업 시작시 자동으로 삭제 후 생성됨)
```

## 프로그램 실행방법

1. nodejs 18.15.0 LTS 설치 (https://nodejs.org/ko)

2. 변환할 데이터를 ./01_input 폴더로 이동

```
 ㄴ 01_input
    ㄴ 59            // 장비번호
        ㄴ 2301_LensAAMain_01(KOR)_BCR3.txt     // 파일들 (확장자 무관, 파일 내용이 csv 형태면 됨)
        ㄴ 2301_LensAAMain_02(KOR)_BCR3.txt
        ㄴ 2301_LensAAMain_03(KOR)_BCR3.txt
```

3. 프로그램 실행

```

$ cd csv2xlsx
$ npm i     // 최초 한번만
$ npm run start

```

4. 결과 파일 확인 ▶ ./02_output
