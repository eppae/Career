function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('checkcount갱신')
        .addItem('데이터 초기화', 'updatedata')
        .addToUi();

    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const targetSheet = spreadsheet.getSheetByName("checkcount");

        if (targetSheet && targetSheet.getLastRow() > 1) {
            Logger.log("'checkcount' 시트에서 데이터 로딩 중");
            loadDataClass(targetSheet);
        } else {
            Logger.log("'checkcount' 시트가 비었거나 없습니다. 데이터를 생성합니다.");
            updatedata();
        }
    } catch (error) {
        Logger.log("Error in onOpen: " + error.message);
    }
}

function mapHeaders(headers) {
    return headers.reduce((map, header, index) => {
        map[header] = index + 1;
        return map;
    }, {});
}

function filterAndMapData(data, headerMap) {
    return data
        .filter(row => 
        row.some(cell => cell !== "") && // 기존 빈 행 제외
        row[headerMap["이름"] - 1] &&   // 이름이 비어있지 않음
        row[headerMap["성별"] - 1] &&   // 성별이 비어있지 않음
        row[headerMap["나이"] - 1]  // 나이가 비어있지 않음
    ) // 빈 행 제외
        .map((row, i) => ({
            row: i + 2, // 실제 행 번호
            name: row[headerMap["이름"] - 1],
            gender: row[headerMap["성별"] - 1],
            birthday: row[headerMap["나이"] - 1],
            checkCount: row[headerMap["checkcount"] - 1] || 0,
        }));
}

function updatedata() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const dataSheet = spreadsheet.getSheetByName("감면여부현황");

        if (!dataSheet) throw new Error("Sheet '감면여부현황' does not exist.");

        let targetSheet = spreadsheet.getSheetByName("checkcount");
        if (!targetSheet) {
            targetSheet = spreadsheet.insertSheet("checkcount");
            Logger.log("Sheet 'checkcount' created.");
        }

        const dataObjects = createDataClass(dataSheet);
        updateCheckcountSheet(targetSheet, dataObjects);

    } catch (error) {
        alertMessage("데이터 초기화 중 오류 발생: " + error.message);
        Logger.log("Error in updatedata: " + error.message);
    }
}

function createDataClass(dataSheet) {
    try {
        const datavalues = dataSheet.getDataRange().getValues();
        if (datavalues.length < 2) throw new Error("'감면여부현황' 시트에 데이터가 부족합니다.");

        const headers = datavalues[1]; // 두 번째 행을 헤더로 가정
        const headerMap = mapHeaders(headers);

        const filteredData = filterAndMapData(datavalues.slice(1), headerMap);

        // 중복 제거
        const seen = new Set();
        const uniqueData = filteredData.filter(item => {
            const key = `${item.name}|${item.gender}|${item.birthday}`;
            if (seen.has(key)) {
                return false; // 중복된 데이터는 제외
            }
            seen.add(key);
            return true; // 중복되지 않은 데이터는 추가
        });

        return uniqueData;
    } catch (error) {
        Logger.log("Error in createDataClass: " + error.message);
        return [];
    }
}


function loadDataClass(targetSheet) {
    try {
        const values = targetSheet.getDataRange().getValues();
        if (values.length === 0) throw new Error("'checkcount' 시트가 비어 있습니다.");

        const headers = values[0];
        const headerMap = mapHeaders(headers);
        return filterAndMapData(values.slice(1), headerMap);
    } catch (error) {
        Logger.log("Error in loadDataClass: " + error.message);
        return [];
    }
}

function updateCheckcountSheet(sheet, dataObjects) {
    try {
        // 기존 데이터 가져오기
        const existingData = sheet.getDataRange().getValues();
        let headers = [];
        let headerMap = {};

        if (existingData.length === 0 || existingData[0].length === 0) {
            // 시트가 비어있다면 헤더를 추가
            headers = ["이름", "성별", "나이", "checkCount"];
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        } else {
            headers = existingData[0]; // 첫 번째 행은 헤더
        }

        // 헤더 맵 생성
        headerMap = mapHeaders(headers);

        // 기존 데이터 확인 및 업데이트
        const updatedRows = [];
        const newRows = [];

        // dataObjects의 두 번째 데이터부터 처리
        dataObjects.slice(1).forEach(obj => {
            const existingRow = existingData.find(row =>
                row[headerMap["이름"] - 1] === obj.name &&
                row[headerMap["성별"] - 1] === obj.gender &&
                row[headerMap["나이"] - 1] === obj.birthday
            );

            if (existingRow) {
                // 기존 데이터 업데이트
                const rowIndex = existingData.indexOf(existingRow) + 1; // 실제 행 번호 (1-based)
                sheet.getRange(rowIndex, headerMap["checkCount"]).setValue(obj.checkCount);
                updatedRows.push(rowIndex);
            } else {
                // 새 데이터 추가
                newRows.push([obj.name, obj.gender, obj.birthday, obj.checkCount]);
            }
        });

        // 새 데이터 추가
        if (newRows.length > 0) {
            const startRow = sheet.getLastRow() + 1;
            sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
        }

        Logger.log(`Updated rows: ${updatedRows.length}, New rows: ${newRows.length}`);
    } catch (error) {
        Logger.log("Error in updateCheckcountSheet: " + error.message);
        throw new Error("데이터 업데이트 중 오류 발생: " + error.message);
    }
}


function onEdit(e) {
    try {
        const sheet = e.source.getActiveSheet();
        const row = e.range.getRow();
        const column = e.range.getColumn();
        const discountsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("감면여부현황");

        if (!discountsheet) throw new Error("Sheet '감면여부현황' does not exist.");

        const headers = discountsheet.getRange(2, 1, 2, discountsheet.getLastColumn()).getValues();
        const headerMap = mapHeaders(headers[0]);

        const DiscountCheckIndex = headerMap["감면여부"];
        if (column !== DiscountCheckIndex) return;

        const name = sheet.getRange(row, headerMap["이름"]).getValue();
        const gender = sheet.getRange(row, headerMap["성별"]).getValue();
        const birthday = sheet.getRange(row, headerMap["나이"]).getValue();
        const condition = sheet.getRange(row, headerMap["감면사유"]).getValue();
        const isChecked = sheet.getRange(row, DiscountCheckIndex).getValue();

        // 먼저 checkCount 시트를 업데이트한 후 데이터를 다시 로드합니다.
        const checkCountData = loadDataClass(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("checkcount"));

        alertMessage("checkCountData: " + JSON.stringify(checkCountData));

        const checkCountObject = checkCountData.find(obj =>
            obj.name === name &&
            obj.gender === gender &&
            obj.birthday === birthday
        );

        alertMessage("name: " + name + ", gender: " + gender + ", age: " + birthday);
        alertMessage("checkCountObject: " + JSON.stringify(checkCountObject));

        if (!checkCountObject) throw new Error("일치하는 데이터가 없습니다.");

        const maxCount = condition === "국가유공자" || condition === "수급-생계" ? 1 : 2;
        const colors = maxCount === 1 ? { 1: "#FF6666" } : { 1: "#FFCC66", 2: "#FF6666" };
        const errorMessage = maxCount === 1
            ? "수급-생계, 국가유공자는 최대 1개까지 감면적용이 가능합니다."
            : "수급-기타, 차상위는 최대 2개까지 감면적용이 가능합니다.";

        handleCheckCondition(isChecked, [checkCountObject], row, DiscountCheckIndex, checkCountData, maxCount, colors, errorMessage);

        // 시트를 업데이트한 후 데이터를 다시 반영
        updateCheckcountSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("checkcount"), checkCountData);

    } catch (error) {
        Logger.log("Error in onEdit: " + error.message);
        alertMessage("데이터 수정 중 오류 발생: " + error.message);
    }
}


function handleCheckCondition(isChecked, matchedObjects, row, DiscountCheckIndex, dataObjects, maxCount, colors, errorMessage) {
    matchedObjects.forEach(obj => {
        let checkCount = obj.checkCount;

        if (isChecked && checkCount < maxCount) {
            checkCount++;
            obj.checkCount = checkCount;
            setRowColor(row, checkCount, colors);
        } else if (isChecked) {
            alertMessage(errorMessage);
            resetRowColor(row);
        } else {
            checkCount = Math.max(0, checkCount - 1);
            obj.checkCount = checkCount;
            checkCount > 0 ? setRowColor(row, checkCount, colors) : resetRowColor(row);
        }
    });
}

function setRowColor(row, checkCount, colors) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("감면여부현황");
    const color = colors[checkCount] || null;
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(color);
}

function resetRowColor(row) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("감면여부현황");
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(null);
}

function alertMessage(message) {
    SpreadsheetApp.getUi().alert(message);
}