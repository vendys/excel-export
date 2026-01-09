// ==== Prod ====
const ExcelJS = require("exceljs");
const FileSaver = require("file-saver");
// ==============

const DEFAULT_FILE_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
const DEFAULT_FILE_EXTENSION = ".xlsx";
const DEFAULT_COL_WIDTH = 9;
const WIDTH_ERROR_MARGIN = 1.3;

const generateSheet = async (workbook, sheetLayout, sheetData) => {
  const {
    workSheet: { name, ...workSheetProps },
    rowLayouts = [],
    colLayouts = [],
    cellDatas = [],
    images = [],
    colValidations: colValidationsInLayout = {},
  } = sheetLayout;

  const {
    startRowNum = Infinity, // Info: 새로운 Row 를 넣기 위한 시작점
    newRows = [], // Info: 새롭게 추가할 Row 값
    customValues = [], // Info: 특정 Cell 을 변경해야 하는 위치 및 값
    insertColNum = Infinity, // Info: 새로운 Col 을 넣기 위한 시작점 (Style 은 삽입 전 Col 의 Style 을 따라감, ex - 5번에 삽입 -> 4번과 5번 사이에 추가됨 -> 스타일은 기존 5번 col 의 스타일을 따라감)
    newCols = [], // Info: 새로게 추가할 Col 의 Row 값 (Table 의 head label 을 변경하기 위함)
    customFormulas = [], // Info: Cell 에 지정할 수식
    tableInfos = [], // Info: 스타일을 지정하는 것이 아닌, 단순히 Table 을 넣기 위한 필드
    removeColumns = [], // Info: 특정 Col 을 삭제하기 위한 값
    colValidations = {}, // Info: 특정 Cell 에 대한 Validation 을 담기 위한 변수 (layout 을 통해 추출한 값을 복사 & 붙여넣기)
  } = sheetData;

  const newRowsLength = newRows.length;
  const generateWithoutLayout = Object.keys(workSheetProps).length < 1;

  const { defaultColWidth } = workSheetProps?.properties || DEFAULT_COL_WIDTH;
  const { sheet: isProtectedSheet, ...protectionOptions } =
    workSheetProps?.sheetProtection || { sheet: false };

  const isMovedCol = (colNum) => colNum >= insertColNum;
  const getMovedColNum = (colNum) =>
    isMovedCol(colNum) ? colNum + newCols.length : colNum;

  // Step 01. workbook 파일 props 지정
  const newWorkSheet = workbook.addWorksheet(name, workSheetProps || {});

  if (tableInfos.length > 0) {
    // Hidden Step. Table 로 값을 삽입하기 위해선 스타일 및 value 가 중복 지정되면 안 되기에 분리하여 적용
    tableInfos.forEach((tableInfo) => {
      newWorkSheet.addTable(tableInfo);
    });
  } else {
    // Step 02. Row 스타일 지정
    rowLayouts.forEach(({ number, style, height }) => {
      const isLayoutArea = !newRowsLength || number < startRowNum;

      if (isLayoutArea) {
        newWorkSheet.getRow(number).style = style;
        newWorkSheet.getRow(number).height = height;
      }
    });

    // Step 03. Cell 기본 스타일 및 정보 입력
    const defaultCells = newRowsLength
      ? cellDatas.filter(({ row }) => row < startRowNum)
      : cellDatas;

    defaultCells.forEach(({ style, value, note, col, row }) => {
      const targetColNum = getMovedColNum(col);

      newWorkSheet.getCell(row, targetColNum).value = value;
      newWorkSheet.getCell(row, targetColNum).style = style;

      if (note) {
        newWorkSheet.getCell(row, targetColNum).note = note;
      }
    });

    // Step 04. 새롭게 추가된 col에 대한 value 및 style 추가
    newCols.forEach(({ rowNum, value }, index) => {
      const targetColNum = insertColNum + index;

      newWorkSheet.getColumn(targetColNum).eachCell(({ address, row }) => {
        // Step 04-(1). insertColNum 기준의 스타일을 적용
        const basedCellStyle =
          defaultCells.find((e) => e.col === insertColNum && e.row === row)
            ?.style || {};
        newWorkSheet.getCell(address).style = { ...basedCellStyle };

        if (rowNum === row) {
          newWorkSheet.getCell(address).value = value;
        }
      });
    });

    // Step 05. Col 스타일 지정
    const newWorkSheetCols = [];
    const newColsLength = newCols.length;

    colLayouts.forEach(({ width, number, letter, ...extraProps }) => {
      // Step 05-(1). 셀 너비가 정상적으로 반영되지 않는 이슈가 존재. 대충 비슷하게 맞추기 위해 임의의 상수 부여
      const changedProps = {
        number,
        width: (width || defaultColWidth) * WIDTH_ERROR_MARGIN,
      };

      if (isMovedCol(number)) {
        if (number === insertColNum) {
          newCols.forEach((col, index) => {
            changedProps.number = number + index;
            newWorkSheetCols.push({ ...extraProps, ...changedProps });
          });
        }
        changedProps.number = number + newColsLength;
      }
      newWorkSheetCols.push({ ...extraProps, ...changedProps });
    });

    newWorkSheet.columns = newWorkSheetCols;

    // Step 06. 데이터 삽입 및 validation, style 지정
    if (newRowsLength) {
      newRows.forEach(async (newRow) => {
        await newWorkSheet.addRow(newRow);
      });

      // Step 06-(1). style 기준이 되는 첫 번째 row 의 Cell 필터
      const tableFirstCells = cellDatas.filter(
        ({ row }) => row === startRowNum
      );

      tableFirstCells.forEach(({ style, col, row }) => {
        const targetColNum = getMovedColNum(col);

        for (let i = 0; i <= newRowsLength; i++) {
          // Step 06-(2). 삽입된 데이터 Cell 에 스타일 및 validation 지정
          const currentValidations =
            colValidations[col] || colValidationsInLayout[col];

          if (currentValidations) {
            newWorkSheet.getCell(row + i, targetColNum).dataValidation =
              currentValidations.dataValidation;
          }

          newWorkSheet.getCell(row + i, targetColNum).style = style;

          if (col === insertColNum) {
            // Step 06-(3). 새로 추가된 cols 에 대한 style 을 "삽입된 col" style 과 동일하게 설정
            // ex) 5번과 6번 사이에 새로운 column 을 추가할 경우, 새롭게 추가되는 column 들은 6 번 style 을 따라감
            newCols.forEach((newCol, index) => {
              const targetStyle = tableFirstCells.find(
                (e) => e.col === insertColNum
              ).style;
              newWorkSheet.getCell(row + i, insertColNum + index).style =
                targetStyle;
            });
          }
        }
      });
    }

    // Step 07. 커스텀 value 삽입
    customValues.forEach(({ address, value, note }) => {
      const { row: startRowNum, col: startColNum } =
        newWorkSheet.getCell(address);
      const targetColNum = getMovedColNum(startColNum);

      newWorkSheet.getCell(startRowNum, targetColNum).value = value;
      if (note) {
        newWorkSheet.getCell(startRowNum, targetColNum).note = note;
      }
    });

    // Step 08. 지정한 col 제거
    removeColumns.forEach(({ startColNum, removeColCount }) => {
      const targetColNum = getMovedColNum(startColNum);
      newWorkSheet.spliceColumns(targetColNum, removeColCount);
    });

    // Step 09. 커스텀 formula 삽입
    customFormulas.forEach(({ startAddress, refAddress, formula }) => {
      const { row: startRowNum, col: startColNum } =
        newWorkSheet.getCell(startAddress);
      const targetColNum = getMovedColNum(startColNum);

      const { row: refRowNum, col: refColNum } =
        newWorkSheet.getCell(refAddress);
      const targetRefColNum = getMovedColNum(refColNum);

      for (let i = 0; i <= newRowsLength; i++) {
        const { address: targetAddress } = newWorkSheet.getCell(
          startRowNum + i,
          targetColNum
        );

        const { address: refAddress } = newWorkSheet.getCell(
          refRowNum + i,
          targetRefColNum
        );

        newWorkSheet.getCell(targetAddress).value = {
          formula: formula(refAddress),
        };
      }
    });

    // Step 09-1. 수식 적용하기 (수식에 다른 sheet 정보가 포함된 경우에는 fillFormula 사용 불가, https://github.com/exceljs/exceljs/issues/1766)
    // customFillFormulas.forEach(({ addressRange, formula, values }) => {
    //   newWorkSheet.fillFormula(addressRange, formula, values);
    // });
  }

  // Step 10. Image 삽입
  images.forEach(({ range, imageUrl }) => {
    if (imageUrl) {
      const extension = imageUrl.split(";")[0].split("/")[1];

      const imageId = workbook.addImage({
        base64: imageUrl,
        extension,
      });

      newWorkSheet.addImage(imageId, range);
    }
  });

  // Step 11. Sheet 암호 설정 (기본 "vendys123!")
  if (isProtectedSheet) {
    const sheetPassword =
      "vendys123!" || prompt(`${name} Sheet의 비밀번호를 입력해주세요.`);

    if (sheetPassword) {
      await newWorkSheet.protect(sheetPassword, protectionOptions);
    }
  }
};

const setWorkbookProperties = (workbook, workbookProps) => {
  const { created, creator, lastModifiedBy, modified, views, properties } =
    workbookProps;
  const { date1904, ...extraProps } = properties;

  const newWorkbookProps = {
    creator,
    created: new Date(created),
    lastModifiedBy,
    modified: new Date(modified),
    views,
    properties: { date1904, ...extraProps },
  };

  Object.assign(workbook, newWorkbookProps);
};

const generateBook = (
  workbook,
  sheetLayout = {},
  sheetData = {} // Info: API를 통해 받아온 데이터 및 커스텀 options
) => {
  return new Promise((resolve, reject) => {
    const sheetLayoutKeys = Object.keys(sheetLayout).sort(); // Warning: 문자형으로 sorting하기 때문에 10개 이상의 sheet가 존재할 경우 비정상적으로 동작함;
    const sheetDataKeys = Object.keys(sheetData);

    try {
      sheetLayoutKeys.length > 0
        ? sheetLayoutKeys.forEach((sheetId) => {
            if (sheetId === "workbook") {
              setWorkbookProperties(workbook, sheetLayout[sheetId]);
            } else {
              generateSheet(
                workbook,
                sheetLayout[sheetId],
                sheetData[sheetId] || {}
              );
            }
          })
        : sheetDataKeys.forEach((sheetId) => {
            generateSheet(
              workbook,
              { workSheet: { name: sheetId } },
              sheetData[sheetId] || {}
            );
          });

      resolve(workbook);
    } catch (error) {
      reject(error);
    }
  });
};

const handleFileExport = async (sheetLayout, sheetData, fileName) => {
  const workbook = new ExcelJS.Workbook();
  const newWorkbook = await generateBook(workbook, sheetLayout, sheetData);

  (newWorkbook as any).xlsx
    .writeBuffer()
    .then((buffer) => {
      const excelData = new Blob([buffer], { type: DEFAULT_FILE_TYPE });
      // ==== Dev ====
      // saveAs(excelData, fileName + DEFAULT_FILE_EXTENSION);
      // ==== Prod ====
      FileSaver.saveAs(excelData, fileName + DEFAULT_FILE_EXTENSION);
    })
    .catch((error) => {
      console.error("Error saving file:", error);
    });
};

// ==== Prod ====
export default handleFileExport;
