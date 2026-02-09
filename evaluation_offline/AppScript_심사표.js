// ⚠️ 권한 승인용 테스트 함수 (배포 전 한 번만 실행하세요!)
function authorizePermissions() {
  // 이 함수를 실행하면 Drive 권한 승인 팝업이 뜹니다
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
    var timestamp = new Date(); // 현재 시간을 한 번만 기록
    
    // ✍️ 서명 데이터를 가져옵니다 (base64 이미지)
    var signatureData = e.parameter.signature || "";
    
    // ✍️ 서명을 Google Drive에 이미지 파일로 저장하고 이미지 URL 생성
    var signatureImageUrl = "";
    if (signatureData) {
      try {
        // base64에서 "data:image/png;base64," 부분 제거
        var base64Data = signatureData.split(',')[1];
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', 
                                     judgeName + '_' + Utilities.formatDate(timestamp, "GMT+9", "yyyyMMdd_HHmmss") + '.png');
        
        // 스프레드시트와 같은 폴더에 저장 (권한 문제 최소화)
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var file = DriveApp.getFileById(spreadsheet.getId()).getParents().next().createFile(blob);
        
        // 파일을 누구나 볼 수 있게 공유 설정
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // 이미지 직접 표시용 URL 생성 (IMAGE 함수에서 사용 가능)
        var fileId = file.getId();
        signatureImageUrl = "https://drive.google.com/uc?export=view&id=" + fileId;
        
      } catch (imgError) {
        Logger.log("이미지 저장 오류: " + imgError.toString());
        signatureImageUrl = "오류: " + imgError.toString();
      }
    }

    // 총 3개의 그룹(행)을 처리하는 반복문입니다.
    for (var i = 1; i <= 3; i++) {
      // 각 그룹의 '단체명'을 가져옵니다.
      var groupName = e.parameter["group" + i + "_name"];
      
      // 단체명(groupX_name)이 비어있지 않은 경우에만 시트에 저장합니다.
      if (groupName) { 
        // 폼에서 각 그룹의 데이터를 이름표로 찾습니다.
        var gubun = e.parameter["group" + i + "_gubun"];
        var rep = e.parameter["group" + i + "_rep"];
        var pos = e.parameter["group" + i + "_pos"];
        
        // 점수를 가져오되, 숫자가 아니거나 비어있으면 0으로 처리합니다.
        var s1 = parseFloat(e.parameter["group" + i + "_score1"]) || 0;
        var s2 = parseFloat(e.parameter["group" + i + "_score2"]) || 0;
        var s3 = parseFloat(e.parameter["group" + i + "_score3"]) || 0;
        
        // '점원'이 직접 합계를 계산합니다
        var total = s1 + s2 + s3;

        // ✍️ 시트에 1행(Row)을 추가합니다. (서명 이미지 URL 포함)
        sheet.appendRow([
          timestamp, 
          judgeName, 
          gubun,
          groupName, 
          rep, 
          pos, 
          s1, 
          s2, 
          s3, 
          total,
          signatureImageUrl  // 이미지 직접 표시용 URL
        ]);
      }
    }
    
    // 폼(HTML)에게 "성공!"이라고 JSON 형태로 알려줍니다.
    return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // 만약 오류가 나면 폼에게 "실패!"와 오류 메시지를 알려줍니다.
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
