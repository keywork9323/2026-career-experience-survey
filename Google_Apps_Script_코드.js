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
  if (!sheet) {
    sheet = ss.insertSheet("참여여부");
  }

  if (sheet.getLastRow() === 0) {
    var headers = [
      "제출일시",
      "단과대학명",
      "학부명",
      "트랙명",
      "조교 정보",
      "교원 정보",
      "참여 가능 일정",
      "타 캠퍼스 이동",
      "비고"
    ];
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("#ffffff");
  }

  var row = [
    new Date().toLocaleString("ko-KR"),
    data.college || "",
    data.department || "",
    data.track || "",
    formatPersonList(data.assistants || []),
    formatPersonList(data.professors || []),
    (data.schedule || []).join(", "),
    data.campusMove || "",
    data.remarks || ""
  ];

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

function formatPersonList(persons) {
  return persons
    .map(function (p) {
      return p.name + " (사번:" + p.id + ", " + p.phone + ", " + p.email + ")";
    })
    .join(" / ");
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
