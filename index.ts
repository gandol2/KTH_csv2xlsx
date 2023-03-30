import fs from "fs";
import csv from "csvtojson";
import path from "path";
import XLSX from "xlsx";
import iconvlite from "iconv-lite";

const 양식파일 = "양식.xlsx";
const PATH_INPUT = "01_input";
const PATH_OUTPUT = "02_output";
const OUTPUT_FILE_NAME_TEMPLATE = "_호기_Lens_bcr3_취합.xlsx";
const B_FOLDER_CLEAR = false;

function readFileSync_encoding(filename: string, encoding: string) {
  var content = fs.readFileSync(filename);
  return iconvlite.decode(content, encoding);
}

//* main
const main = async () => {
  console.log(
    "\n\n[START] ==================================================\n"
  );

  if (B_FOLDER_CLEAR) {
    //* 출력 폴더 정리
    try {
      fs.rmSync(path.join(PATH_OUTPUT), { recursive: true, force: true });
      fs.mkdirSync(path.join(PATH_OUTPUT));
    } catch (e) {
      console.log(`[오류 ❌] 출력폴더 정리 실패 (./${PATH_OUTPUT})`);
      throw e;
    }
  }

  const dirList = await fs.readdirSync(PATH_INPUT);
  var allFileCount = 0;
  for (var dirName of dirList) {
    const fileList = await fs.readdirSync(path.join(PATH_INPUT, dirName));
    allFileCount += fileList.length;
    console.log(fileList.length);
  }

  //* 폴더 목록
  var fileCount = 1;
  for (var dirName of dirList) {
    const resultXLSXFileName = `${dirName}${OUTPUT_FILE_NAME_TEMPLATE}`;
    const resultXLSXFilePath = path.join(PATH_OUTPUT, resultXLSXFileName);
    // fs.copyFileSync(양식파일, resultXLSXFilePath);

    console.log(`[🚩작업시작] 폴더명 ▶ ${dirName}`);
    const fileList = await fs.readdirSync(path.join(PATH_INPUT, dirName));
    allFileCount += fileList.length;
    const wb = XLSX.readFile(양식파일, { cellStyles: true });
    const ws_orignal = wb.Sheets["원본"];

    var sheetCount = 1;
    //* 파일 목록
    for (var fileName of fileList) {
      var new_ws = JSON.parse(JSON.stringify(ws_orignal)); // make a copy of the object

      const sheetName = fileName
        .split("_")
        .map((ele, idx) => (idx % 2 === 0 ? ele.replace(/[^0-9]/g, "") : ""))
        .join("");

      const filePath = path.join(PATH_INPUT, dirName, fileName);
      //   console.log(`파일명 : ${fileName}`);
      console.log(
        `[${sheetCount++}/${
          fileList.length
        }] ✔ Sheet 만드는 중... ▶ ${sheetName}`
      );
      const str = readFileSync_encoding(filePath, "euc-kr");

      //   //* 파일 to Array
      const jsonArray = await csv({
        noheader: true,
        output: "csv",
      }).fromString(str);

      jsonArray.shift();

      var bFind_미사용 = false;
      const jsonArray2 = jsonArray.map((x: string[]) => {
        if (0 < x.length && !bFind_미사용 && x[0].includes("미사용")) {
          bFind_미사용 = true;
        }

        if (bFind_미사용) {
          x.unshift("");
        }

        return x;
      });
      //   console.log(jsonArray2);

      //   console.log(jsonArray[0]);
      XLSX.utils.book_append_sheet(wb, new_ws, sheetName); // append the worksheet
      const ws = wb.Sheets[sheetName];
      XLSX.utils.sheet_add_aoa(ws, jsonArray, { origin: "E3" });
    }

    console.log(
      `[${fileCount++}/${
        dirList.length
      }] ✅ 파일 저장중... ▶ ${resultXLSXFilePath}`
    );
    console.log("");
    XLSX.writeFileXLSX(wb, resultXLSXFilePath, { cellStyles: false });
  }

  //   dirList.forEach(async (dirName) => {
  //     console.log(dirName);
  //     const fileList = await fs.readdirSync(path.join(PATH_INPUT, dirName));
  //     console.log(fileList);

  //     fileList.forEach(async (fileName) => {
  //       console.log(fileName);
  //     });
  //   });

  //* Input 폴더 초기화
  if (B_FOLDER_CLEAR) {
    try {
      fs.rmSync(path.join(PATH_INPUT), { recursive: true, force: true });
      fs.mkdirSync(path.join(PATH_INPUT));
    } catch (e) {
      console.log(`[오류 ❌] 입력폴더 정리 실패 (./${PATH_OUTPUT})`);
      throw e;
    }
  }

  console.log(`[전체 입력 파일 수] ${allFileCount.toLocaleString()} 개`);

  console.log(
    "\n[END] ----------------------------------------------------\n\n"
  );
};

//* top level async
(async () => {
  try {
    const text = await main();
    console.log("[완료 ✅]");
  } catch (e) {
    console.log("[에러 ❌]");
    console.log(e);
  }
})();
