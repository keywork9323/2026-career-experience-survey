/**
 * ============================================================
 * 2026 고교생 진로체험 프로그램 - Google Apps Script
 * ============================================================
 *
 * [설정 방법]
 * 1. Google 스프레드시트를 새로 만듭니다.
 * 2. 시트 이름을 2개 만듭니다:
 *    - "참여여부" (첫 번째 시트)
 *    - "운영계획서" (두 번째 시트)
 * 3. 상단 메뉴 → 확장 프로그램 → Apps Script 클릭
 * 4. 기본 코드를 모두 지우고, 이 파일의 내용을 붙여넣기 합니다.
 * 5. 상단 메뉴 → 배포 → 새 배포
 *    - 유형: 웹 앱
 *    - 실행 주체: 나
 *    - 액세스 권한: 모든 사용자
 * 6. 배포 후 나오는 URL을 복사합니다.
 * 7. 두 HTML 파일에서 SCRIPT_URL 변수에 해당 URL을 붙여넣습니다.
 * ============================================================
 */

var SPREADSHEET_ID = '1JDWkiU_ug0jIPJS-ZJ8yCTuU8k_u-ppJ8f3hPiXNyO8';

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    result: 'success',
    message: '2026 진로체험 프로그램 설문 API가 정상 작동 중입니다.'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var formType = data.formType;

    if (formType === "participation") {
      return handleParticipation(ss, data);
    } else if (formType === "plan") {
      return handlePlan(ss, data);
    }

    return createResponse({ result: "error", message: "Unknown form type" });
  } catch (error) {
    return createResponse({ result: "error", message: error.toString() });
  }
}

function handleParticipation(ss, data) {
  var sheet = ss.getSheetByName("참여여부");
  if (!sheet) { sheet = ss.insertSheet("참여여부"); }

  var assistantCount = parseInt(data.assistant_count) || 0;
  var professorCount = parseInt(data.professor_count) || 0;
  var scheduleCount = parseInt(data.schedule_count) || 0;

  // ── 헤더 생성 헬퍼 ──
  function buildHeaders(maxA, maxP, maxS) {
    var h = ["제출일시", "단과대학명", "학부명", "트랙명"];
    for (var i = 1; i <= maxA; i++) {
      h.push("조교" + i + "_이름", "조교" + i + "_사번", "조교" + i + "_연락처", "조교" + i + "_이메일");
    }
    for (var i = 1; i <= maxP; i++) {
      h.push("교원" + i + "_이름", "교원" + i + "_사번", "교원" + i + "_연락처", "교원" + i + "_이메일");
    }
    for (var i = 1; i <= maxS; i++) {
      h.push("일정" + i);
    }
    h.push("타 캠퍼스 이동", "비고");
    return h;
  }

  // ── 헤더 스타일 적용 ──
  function styleHeaders(sh, colCount) {
    var range = sh.getRange(1, 1, 1, colCount);
    range.setFontWeight("bold");
    range.setBackground("#4285f4");
    range.setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }

  // ── 헤더에서 현재 최대 번호 파악 ──
  function getMaxFromHeaders(headers, prefix, suffix) {
    var max = 0;
    var regex = new RegExp("^" + prefix + "(\\d+)" + (suffix || "_이름") + "$");
    for (var i = 0; i < headers.length; i++) {
      var m = headers[i].toString().match(regex);
      if (m) max = Math.max(max, parseInt(m[1]));
    }
    return max;
  }

  // ── 기존 데이터에서 실제 사용된 최대 번호 파악 ──
  function getActualMaxFromData(allData, headers, prefix, suffix) {
    var max = 0;
    var regex = new RegExp("^" + prefix + "(\\d+)" + (suffix || "_이름") + "$");
    for (var r = 0; r < allData.length; r++) {
      for (var c = 0; c < headers.length; c++) {
        var m = headers[c].toString().match(regex);
        if (m && allData[r][c] && allData[r][c].toString().trim() !== "") {
          max = Math.max(max, parseInt(m[1]));
        }
      }
    }
    return max;
  }

  if (sheet.getLastRow() === 0) {
    // ── 신규 시트: 제출 데이터 기준으로 헤더 생성 ──
    var maxA = Math.max(assistantCount, 1);
    var maxP = Math.max(professorCount, 1);
    var maxS = Math.max(scheduleCount, 1);
    var headers = buildHeaders(maxA, maxP, maxS);
    sheet.appendRow(headers);
    styleHeaders(sheet, headers.length);
  } else {
    // ── 기존 시트: 필요 시 열 확장/축소 ──
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var currentMaxA = getMaxFromHeaders(existingHeaders, "조교", "_이름");
    var currentMaxP = getMaxFromHeaders(existingHeaders, "교원", "_이름");
    var currentMaxS = getMaxFromHeaders(existingHeaders, "일정", "$");

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var allData = [];
    if (lastRow > 1) {
      allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    }

    // 실제 데이터가 있는 최대 수 + 이번 제출 수 중 큰 값
    var actualMaxA = getActualMaxFromData(allData, existingHeaders, "조교", "_이름");
    var actualMaxP = getActualMaxFromData(allData, existingHeaders, "교원", "_이름");
    var actualMaxS = getActualMaxFromData(allData, existingHeaders, "일정", "$");
    var neededMaxA = Math.max(assistantCount, actualMaxA, 1);
    var neededMaxP = Math.max(professorCount, actualMaxP, 1);
    var neededMaxS = Math.max(scheduleCount, actualMaxS, 1);

    if (neededMaxA !== currentMaxA || neededMaxP !== currentMaxP || neededMaxS !== currentMaxS) {
      // 헤더 매핑 생성
      var oldMap = {};
      for (var i = 0; i < existingHeaders.length; i++) {
        oldMap[existingHeaders[i]] = i;
      }
      var newHeaders = buildHeaders(neededMaxA, neededMaxP, neededMaxS);
      var newMap = {};
      for (var i = 0; i < newHeaders.length; i++) {
        newMap[newHeaders[i]] = i;
      }

      // 기존 데이터를 새 구조에 맞게 재배치
      var newData = [];
      for (var r = 0; r < allData.length; r++) {
        var newRow = new Array(newHeaders.length).fill("");
        for (var h in oldMap) {
          if (newMap.hasOwnProperty(h)) {
            newRow[newMap[h]] = allData[r][oldMap[h]];
          }
        }
        newData.push(newRow);
      }

      // 시트 재작성
      sheet.clear();
      sheet.appendRow(newHeaders);
      if (newData.length > 0) {
        sheet.getRange(2, 1, newData.length, newHeaders.length).setValues(newData);
      }
      styleHeaders(sheet, newHeaders.length);
    }
  }

  // ── 현재 최종 헤더 기준으로 데이터 행 작성 ──
  var finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var finalMaxA = getMaxFromHeaders(finalHeaders, "조교", "_이름");
  var finalMaxP = getMaxFromHeaders(finalHeaders, "교원", "_이름");
  var finalMaxS = getMaxFromHeaders(finalHeaders, "일정", "$");

  var row = [
    new Date().toLocaleString("ko-KR"),
    data.college || "",
    data.department || "",
    data.track || ""
  ];

  for (var i = 1; i <= finalMaxA; i++) {
    row.push(
      i <= assistantCount ? (data["assistant_name_" + i] || "") : "",
      i <= assistantCount ? (data["assistant_id_" + i] || "") : "",
      i <= assistantCount ? (data["assistant_phone_" + i] || "") : "",
      i <= assistantCount ? (data["assistant_email_" + i] || "") : ""
    );
  }

  for (var i = 1; i <= finalMaxP; i++) {
    row.push(
      i <= professorCount ? (data["professor_name_" + i] || "") : "",
      i <= professorCount ? (data["professor_id_" + i] || "") : "",
      i <= professorCount ? (data["professor_phone_" + i] || "") : "",
      i <= professorCount ? (data["professor_email_" + i] || "") : ""
    );
  }

  for (var i = 1; i <= finalMaxS; i++) {
    row.push(i <= scheduleCount ? (data["schedule_" + i] || "") : "");
  }

  row.push(
    data.campusMove || "",
    data.remarks || ""
  );

  sheet.appendRow(row);
  return createResponse({ result: "success", message: "참여 여부가 저장되었습니다." });
}

function handlePlan(ss, data) {
  var sheet = ss.getSheetByName("운영계획서");
  if (!sheet) {
    sheet = ss.insertSheet("운영계획서");
  }

  if (sheet.getLastRow() === 0) {
    var headers = [
      "제출일시",
      "트랙명",
      "교원명",
      "프로그램 목표",
      "트랙 소개",
      "전공특강 주제",
      "전공특강 내용",
      "체험활동 주제",
      "체험활동 내용"
    ];
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("#ffffff");
  }

  var row = [
    new Date().toLocaleString("ko-KR"),
    data.trackName || "",
    data.professorName || "",
    data.programGoal || "",
    data.trackIntro || "",
    data.lectureTopic || "",
    data.lectureContent || "",
    data.activityTopic || "",
    data.activityContent || ""
  ];

  sheet.appendRow(row);
  return createResponse({ result: "success", message: "운영 계획서가 저장되었습니다." });
}


function createResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// 권한 승인용 - 편집기에서 이 함수를 한 번 실행해 주세요
function authorize() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("스프레드시트 이름: " + ss.getName());
  Logger.log("권한 승인 완료!");
}
