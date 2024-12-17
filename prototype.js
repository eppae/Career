// 전역 변수 선언
let dataObjects = [];

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('checkcount갱신') // 메뉴 이름
        .addItem('데이터 초기화', 'updatedata') // 메뉴 항목 추가
        .addToUi(); // UI에 추가

    // 스프레드시트를 열 때 데이터 로드
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = spreadsheet.getSheetByName("checkcount");

    if (targetSheet && targetSheet.getLastRow() > 1) {
        dataObjects = loadDataClass(targetSheet);
        Logger.log("dataObjects loaded from 'checkcount'");
    } else {
        Logger.log("'checkcount' sheet is empty or does not exist.");
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
        .filter(row => row.some(cell => cell !== "")) // 빈 행 제외
        .map((row, i) => ({
            row: i + 2, // 실제 행 번호
            name: row[headerMap["이름"] - 1],
            gender: row[headerMap["성별"] - 1],
            birthday: row[headerMap["생년월일"] - 1],
            checkCount: row[headerMap["checkcount"] - 1] || 0,
        }));
}

function updatedata() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const dataSheet = spreadsheet.getSheetByName("감면여부현황");
    if (!dataSheet) {
        Logger.log("Sheet '감면여부현황' does not exist.");
        return;
    }

    let targetSheet = spreadsheet.getSheetByName("checkcount");
    if (!targetSheet) {
        targetSheet = spreadsheet.insertSheet("checkcount");
        Logger.log("Sheet 'checkcount' created.");
        const newdataObjects = createDataClass(dataSheet); // 새로운 데이터 생성
        updateCheckcountSheet(targetSheet, newdataObjects); // 갱신
    } else {
        Logger.log("Sheet 'checkcount' already exists.");
        const loaddataObjects = loadDataClass(targetSheet);
        updateCheckcountSheet(targetSheet, loaddataObjects)
    }
}

function createDataClass(dataSheet) {
    const datavalues = dataSheet.getDataRange().getValues();
    if (datavalues.length === 0) {
        Logger.log("'감면여부현황' 시트의 내용이 존재하지 않습니다.");
        return [];
    }

    const headers = datavalues[1]; //감면여부 현황은 두 번째 행을 헤더로 사용
    const headerMap = mapHeaders(headers);

    const dataObjects = filterAndMapData(datavalues.slice(1), headerMap)

    Logger.log(dataObjects);
    return dataObjects;
}

function loadDataClass(targetSheet) {
    const values = targetSheet.getDataRange().getValues();
    if (values.length === 0) {
        Logger.log("'checkcount' 시트의 내용이 존재하지 않습니다.");
        return [];
    }

    const headers = values[0]; // 첫 번째 행을 헤더로 사용
    const headerMap = mapHeaders(headers);

    const dataObjects = filterAndMapData(values.slice(1), headerMap)

    Logger.log("Loaded dataObjects: ", dataObjects);
    return dataObjects;
}

function updateCheckcountSheet(sheet, dataObjects) {
    sheet.clear();
    sheet.appendRow(["이름", "성별", "생년월일", "checkcount"]); // 헤더 추가

    dataObjects.forEach(obj => {
        sheet.appendRow([obj.name, obj.gender, obj.birthday, obj.checkCount]);
    });

    Logger.log("'checkcount' sheet updated successfully.");
}


// 편집 이벤트 처리 (onEdit)
function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const column = range.getColumn();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const discountsheet = spreadsheet.getSheetByName("감면여부현황");

    if (!discountsheet) {
        Logger.log("감면여부현황 시트가 없습니다");
        return;
    }

    const currentvalue = discountsheet.getDataRange().getValues();
    const currentheaders = currentvalue[1]; // 헤더는 두 번째 행
    const currentheaderMap = mapHeaders(currentheaders);

    const NameIndex = currentheaderMap["이름"];
    const GenderIndex = currentheaderMap["성별"];
    const BirthIndex = currentheaderMap["생년월일"];
    const DiscountCheckIndex = currentheaderMap["감면여부"];
    const Condition = currentheaderMap["감면사유"];
    const ProgramEnd = currentheaderMap["프로그램종결여부"];
    const Percent = currentheaderMap["감면금액비율"];

    const dataRange = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const name = dataRange[NameIndex -1];
    const gender = dataRange[GenderIndex -1];
    const birth = dataRange[BirthIndex -1];
     

    let targetSheet = spreadsheet.getSheetByName("checkcount");
    if (!targetSheet) {
        Logger.log("checkcount 시트가 없습니다");
        return;
    }

    const targetvalues = targetSheet.getDataRange().getValues();
    const targetheaders = targetvalues[0]; // 헤더는 첫 번째 행
    const targetheaderMap = mapHeaders(targetheaders);

    let dataObjects = filterAndMapData(targetvalues.slice(1), targetheaderMap);



    
    const dbNameIndex = targetheaderMap["이름"];
    const dbcheckCountIndex = targetheaderMap["checkcount"];
    const dbGenderIndex = targetheaderMap["성별"];
    const dbBirthdayIndex = targetheaderMap["생년월일"];

    const targetRow = targetvalues.findIndex(row => row[dbNameIndex -1] === name && row[dbGenderIndex -1] === gender && row[dbBirthdayIndex -1]);

    // 감면여부 열이 수정된 경우에만 실행
    if (column === DiscountCheckIndex) {
        // 현재 행의 데이터 가져오기
        const name = dataRange[NameIndex - 1];
        const gender = dataRange[GenderIndex - 1];
        const birth = dataRange[BirthIndex - 1];
        const condition = dataRange[Condition - 1];

        // dataObjects 다시 로드 (최신 상태 유지)
        const targetvalues = targetSheet.getDataRange().getValues();
        const targetheaders = targetvalues[0];
        const targetheaderMap = mapHeaders(targetheaders);
        let dataObjects = filterAndMapData(targetvalues.slice(1), targetheaderMap);

        // 현재 체크 상태
        const isChecked = sheet.getRange(row, DiscountCheckIndex).getValue();


        // 현재 name, gender, birth와 일치하는 객체 찾기
        const matchedObjects = dataObjects.filter(obj => 
            obj.name === name && 
            obj.gender === gender && 
            obj.birthday === birth
        );

        if (matchedObjects.length === 0){
            Logger.log("해당 사용자의 데이터가 존재하지 않습니다.")
        }

        // 체크카운트 계산 --> 이거 ++ 하는 기능이랑 같이 함수로 빼서 작성할 것임 function checkcountplus, checkcountminus
        const checkCount = matchedObjects.checkCount;

        // 조건에 따른 처리
        let shouldRevert = false;
        switch (condition) {
            case "국가유공자":
            case "수급-생계":
                shouldRevert = handleCheck100Condition(checkCount);
                break;
            case "수급-기타":
            case "차상위":
                shouldRevert = handleCheck50Condition(checkCount);
                break;
            case "형제자매":
                // 추가 조건 처리
                break;
            default:
                Logger.log("처리되지 않은 조건:", condition);
        }

        // 조건 위반 시 되돌리기
        if (shouldRevert) {
            sheet.getRange(row, DiscountCheckIndex).setValue(e.oldValue);
            return;
        }

        // dataObjects 업데이트
        matchedObjects.forEach(obj => {
            obj.checkCount = checkCount;
        });

        // checkcount 시트 업데이트
        updateCheckcountSheet(targetSheet, dataObjects);
    }
}

// 100% 감면 조건 처리
function handleCheck100Condition(checkCount) {
    // 추가 제한 조건 확인
    // 예: 최대 체크 가능 횟수 초과 시 true 반환
    if (checkCount > 1) {
        return true; // 되돌리기
    }
    return false; // 진행 허용
}

// 50% 감면 조건 처리
function handleCheck50Condition(checkCount) {
    // 추가 제한 조건 확인
    // 예: 최대 체크 가능 횟수 초과 시 true 반환
    if (checkCount > 2) {
        return true; // 되돌리기
    }
    return false; // 진행 허용
}