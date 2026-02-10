// ⚠️ 권한 승인용 테스트 함수 (배포 전 한 번만 실행하세요!)
function authorizePermissions() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();
  Logger.log("권한 승인 완료! 폴더: " + folder.getName());
}

// doPost 함수는 HTML 폼에서 'POST' 요청(제출)을 받을 때 실행됩니다.
function doPost(e) {
  try {
    // 'Scores'라는 이름의 시트를 찾습니다.
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scores");

    // 폼에서 '심사위원' 이름을 가져옵니다.
    var judgeName = e.parameter.judgeName || "익명";
    var timestamp = new Date();

    // ✍️ 서명 데이터를 가져옵니다 (base64 이미지)
    var signatureData = e.parameter.signature || "";

    // ✍️ 서명을 Google Drive에 이미지 파일로 저장하고 URL 생성
    var signatureImageUrl = "";
    if (signatureData) {
      try {
        var base64Data = signatureData.split(',')[1];
        var blob = Utilities.newBlob(
          Utilities.base64Decode(base64Data),
          'image/png',
          judgeName + '_' + Utilities.formatDate(timestamp, "GMT+9", "yyyyMMdd_HHmmss") + '.png'
        );

        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var file = DriveApp.getFileById(spreadsheet.getId()).getParents().next().createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        var fileId = file.getId();
        signatureImageUrl = "https://drive.google.com/uc?export=view&id=" + fileId;

      } catch (imgError) {
        Logger.log("이미지 저장 오류: " + imgError.toString());
        signatureImageUrl = "오류: " + imgError.toString();
      }
    }

    // ================================
    // 기관별 설정 정의
    // ================================
    var institutions = [
      {
        prefix: "keti",
        name: "한국전자기술연구원",
        maxScores: [25, 35, 40, 0],   // score4는 해당없음이므로 0
        hasScore4: false               // 취업 연계 항목 없음
      },
      {
        prefix: "inha",
        name: "인하대학교",
        maxScores: [25, 35, 25, 15],
        hasScore4: true
      },
      {
        prefix: "kau",
        name: "한국항공대학교",
        maxScores: [25, 35, 25, 15],
        hasScore4: true
      },
      {
        prefix: "kwu",
        name: "광운대학교",
        maxScores: [25, 35, 25, 15],
        hasScore4: true
      }
    ];

    // ================================
    // 기관별로 시트에 한 행씩 기록
    // ================================
    for (var i = 0; i < institutions.length; i++) {
      var inst = institutions[i];
      var prefix = inst.prefix;

      var s1 = parseFloat(e.parameter[prefix + "_score1"]) || 0;
      var s2 = parseFloat(e.parameter[prefix + "_score2"]) || 0;
      var s3 = parseFloat(e.parameter[prefix + "_score3"]) || 0;
      // 한국전자기술연구원은 score4가 hidden(0)으로 넘어오므로 그대로 0 처리
      var s4 = inst.hasScore4 ? (parseFloat(e.parameter[prefix + "_score4"]) || 0) : 0;

      var total = s1 + s2 + s3 + s4;

      sheet.appendRow([
        timestamp,                          // 제출 시각
        judgeName,                          // 심사위원 성명
        inst.name,                          // 기관명
        s1,                                 // 성과 목표 달성 수준 (25점)
        s2,                                 // 성과관리 및 환류 체계의 우수성 (35점)
        s3,                                 // 계획 수립의 적절성 (keti:40점 / 대학:25점)
        inst.hasScore4 ? s4 : "해당없음",   // 취업 연계 및 성과 확산의 적절성
        total,                              // 합계 (100점)
        signatureImageUrl                   // 서명 이미지 URL
      ]);
    }

    // 성공 응답
    return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // 오류 응답
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
