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

        const headers = datavalues[1];
        const headerMap = mapHeaders(headers);
        return filterAndMapData(datavalues.slice(1), headerMap);
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
    sheet.clear(); // 시트 전체 초기화

    // 헤더를 명시적으로 추가
    // 데이터 추가 (헤더와 동일한 데이터를 추가하지 않음)
    dataObjects.forEach(obj => {
        // 모든 필드가 유효한 경우에만 추가
        if (obj.name && obj.gender && obj.birthday) {
            sheet.appendRow([obj.name, obj.gender, obj.birthday, obj.checkCount]);
        }
    });

    Logger.log("'checkcount' sheet updated successfully.");
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

        const checkCountData = loadDataClass(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("checkcount"));
        const checkCountObject = checkCountData.find(obj =>
            obj.name === name &&
            obj.gender === gender &&
            obj.birthday === birthday
        );

        
        // alertMessage("checkCountObject: " + JSON.stringify(checkCountObject));      
        //alertMessage(`name: ${name}, gender: ${gender}, birthday: ${birthday}`);

        if (!checkCountObject) throw new Error("일치하는 데이터가 없습니다.");

        const maxCount = condition === "국가유공자" || condition === "수급-생계" ? 1 : 2;
        const colors = maxCount === 1 ? { 1: "#FF6666" } : { 1: "#FFCC66", 2: "#FF6666" };
        const errorMessage = maxCount === 1
            ? "수급-생계, 국가유공자는 최대 1개까지 감면적용이 가능합니다."
            : "수급-기타, 차상위는 최대 2개까지 감면적용이 가능합니다.";

        handleCheckCondition(isChecked, [checkCountObject], row, DiscountCheckIndex, checkCountData, maxCount, colors, errorMessage);
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
