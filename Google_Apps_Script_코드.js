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

  // ── 고정 일정 목록 (38개) ──
  var FIXED_DATES = [
    "2026.5.6.(수)", "2026.5.7.(목)", "2026.5.8.(금)",
    "2026.5.12.(화)", "2026.5.13.(수)", "2026.5.14.(목)", "2026.5.15.(금)",
    "2026.5.19.(화)", "2026.5.20.(수)", "2026.5.21.(목)", "2026.5.22.(금)",
    "2026.5.26.(화)", "2026.5.27.(수)", "2026.5.28.(목)", "2026.5.29.(금)",
    "2026.6.23.(화)", "2026.6.24.(수)", "2026.6.25.(목)", "2026.6.26.(금)", "2026.6.30.(화)",
    "2026.7.7.(화)", "2026.7.8.(수)", "2026.7.9.(목)", "2026.7.10.(금)",
    "2026.7.14.(화)", "2026.7.15.(수)", "2026.7.16.(목)", "2026.7.17.(금)",
    "2026.7.21.(화)", "2026.7.22.(수)",
    "2026.8.4.(화)", "2026.8.5.(수)", "2026.8.6.(목)", "2026.8.7.(금)",
    "2026.8.11.(화)", "2026.8.12.(수)", "2026.8.13.(목)", "2026.8.14.(금)"
  ];

  // ── 헤더 생성 헬퍼 ──
  function buildHeaders(maxA, maxP) {
    var h = ["제출일시", "단과대학명", "학부명", "트랙명"];
    for (var i = 1; i <= maxA; i++) {
      h.push("조교" + i + "_이름", "조교" + i + "_사번", "조교" + i + "_연락처", "조교" + i + "_이메일", "조교" + i + "_내선번호");
    }
    for (var i = 1; i <= maxP; i++) {
      h.push("교원" + i + "_이름", "교원" + i + "_사번", "교원" + i + "_연락처", "교원" + i + "_이메일", "교원" + i + "_내선번호");
    }
    for (var i = 0; i < FIXED_DATES.length; i++) {
      h.push(FIXED_DATES[i]);
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
    var headers = buildHeaders(maxA, maxP);
    sheet.appendRow(headers);
    styleHeaders(sheet, headers.length);
  } else {
    // ── 기존 시트: 필요 시 조교/교원 열 확장/축소 ──
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var currentMaxA = getMaxFromHeaders(existingHeaders, "조교", "_이름");
    var currentMaxP = getMaxFromHeaders(existingHeaders, "교원", "_이름");

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var allData = [];
    if (lastRow > 1) {
      allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    }

    var actualMaxA = getActualMaxFromData(allData, existingHeaders, "조교", "_이름");
    var actualMaxP = getActualMaxFromData(allData, existingHeaders, "교원", "_이름");
    var neededMaxA = Math.max(assistantCount, actualMaxA, 1);
    var neededMaxP = Math.max(professorCount, actualMaxP, 1);

    if (neededMaxA !== currentMaxA || neededMaxP !== currentMaxP) {
      var oldMap = {};
      for (var i = 0; i < existingHeaders.length; i++) {
        oldMap[existingHeaders[i]] = i;
      }
      var newHeaders = buildHeaders(neededMaxA, neededMaxP);
      var newMap = {};
      for (var i = 0; i < newHeaders.length; i++) {
        newMap[newHeaders[i]] = i;
      }

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
      i <= assistantCount ? (data["assistant_email_" + i] || "") : "",
      i <= assistantCount ? (data["assistant_ext_" + i] || "") : ""
    );
  }

  for (var i = 1; i <= finalMaxP; i++) {
    row.push(
      i <= professorCount ? (data["professor_name_" + i] || "") : "",
      i <= professorCount ? (data["professor_id_" + i] || "") : "",
      i <= professorCount ? (data["professor_phone_" + i] || "") : "",
      i <= professorCount ? (data["professor_email_" + i] || "") : "",
      i <= professorCount ? (data["professor_ext_" + i] || "") : ""
    );
  }

  // ── 고정 일정 열: 선택된 날짜는 O, 아니면 빈칸 ──
  for (var i = 0; i < FIXED_DATES.length; i++) {
    row.push(data[FIXED_DATES[i]] === "O" ? "O" : "");
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
