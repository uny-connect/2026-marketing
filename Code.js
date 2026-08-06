/**
 * ===============================================================================
 * [메인 제어 타워] 통합 메뉴 및 실시간 이벤트 핸들러
 * ===============================================================================
 */

/******************************************************************************************************************
 * [공통] 시트 상단 메뉴 생성 (2026 운영 봇 + 2027 관리자/보고서 메뉴 완전 통합)
 *****************************************************************************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🤖 취재 알람 등록봇')
    .addItem('기존 취재 캘린더에 보내기', 'createCoverageSchedules')
    .addToUi();

  ui.createMenu('📅 일본 취재 등록봇')
    .addItem('지금 즉시 일정 등록', 'checkAndRegisterSchedules')
    .addItem('매일 새벽 자동 실행 설정', 'createJapanTripTrigger')
    .addToUi();

  ui.createMenu('🚀 보고서 자동화')
    .addItem('업체별 시트 동기화 (행 지정)', 'runSyncWithPrompt')
    .addToUi();

  ui.createMenu('⚙️ 관리자 메뉴')
    .addItem('1.새 일정 블록(템플릿) 추가', 'copyTemplateBlock')
    .addItem('2.선택한 셀에 접속 코드 발급', 'generateCodesForSelection')
    .addItem('3.D열 자동 프로필 감지 센서 구동', 'setupProfileTrigger')
    .addItem('4.선택한 행 블로그 타이틀 자동 추출/번역', 'fetchAndTranslateBlogTitles')
    .addToUi();
}

/******************************************************************************************************************
 * [공통 이벤트 핸들러] 프로젝트 내 유일한 실시간 편집 감지 라우터
 *****************************************************************************************************************/
function onEdit(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  // 1. 스케줄 관련 시트인 경우 -> D열 네이버 프로필 주소 자동 감지 센서 작동
  if (sheetName.includes("Schedule") || sheetName.includes("スケジュール")) {
    onEditBloggerUrl(e);
  }
  
  // 2. 보고서 시트인 경우 (미사용 기능 - 안전을 위해 선언 여부 검사 후 실행)
  if (sheetName === CONFIG.SHEETS.REPORT) {
    if (typeof onEditAutoCommentTranslate === 'function') {
      onEditAutoCommentTranslate(e);
    }
  }
}

/**
 * ===============================================================================
 * [관리자 전용 기능 및 유틸리티]
 * ===============================================================================
 */

/**
 * [관리자 기능] 템플릿 영역(2~64행) 맨 아래로 복사 및 행 높이 강제 고정
 */
function copyTemplateBlock() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const templateRange = sheet.getRange(2, 1, 63, sheet.getLastColumn());
  const targetRow = sheet.getLastRow() + 1;
  const targetRange = sheet.getRange(targetRow, 1, 63, sheet.getLastColumn());
  
  templateRange.copyTo(targetRange);
  SpreadsheetApp.flush(); 
  
  targetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
  for (let i = 0; i < 63; i++) {
    sheet.setRowHeight(targetRow + i, 21);
  }
  
  SpreadsheetApp.getUi().alert('✅ 완료');
}

/******************************************************************************************************************
 * [관리자 기능] 고유 코드 생성
 *****************************************************************************************************************/
function generateCodesForSelection() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789@#';
  let codesGenerated = 0;
  
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] === "") {
        let code = '';
        for (let k = 0; k < 6; k++) {
          code += chars.charAt(Math.floor(Math.random() * chars.length));
        }
        values[i][j] = code; 
        codesGenerated++;
      }
    }
  }
  
  if (codesGenerated > 0) {
    range.setValues(values);
  } else {
    SpreadsheetApp.getUi().alert('⚠️ 빈 셀이 아닙니다.');
  }
}

/******************************************************************************************************************
 * [관리자 기능] D열 자동 프로필 감지 센서 구동 및 트리거 등록
 *****************************************************************************************************************/
function setupProfileTrigger() { 
  const triggers = ScriptApp.getProjectTriggers(); 
  for (let i = 0; i < triggers.length; i++) { 
    if (triggers[i].getHandlerFunction() === 'onEditBloggerUrl') { 
      ScriptApp.deleteTrigger(triggers[i]); 
    } 
  } 
  ScriptApp.newTrigger('onEditBloggerUrl').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create(); 
  SpreadsheetApp.getUi().alert("✅ 2026년용 D열 자동 프로필 감지 센서가 완벽하게 설치되었습니다!"); 
}

/******************************************************************************************************************
 * [내부 로직] D열 네이버 프로필 주소 자동 감지 상세 처리
 *****************************************************************************************************************/
function onEditBloggerUrl(e) { 
  const sheet = e.source.getActiveSheet(); 
  const range = e.range; 
  const editedCol = range.getColumn(); 
  const editedRow = range.getRow(); 
  
  if (editedCol === 4 && editedRow > 2) { 
    const blogUrl = range.getValue().toString().trim(); 
    const profileCell = sheet.getRange(editedRow, 6); 
    
    if (blogUrl === "") { 
      profileCell.setValue(""); 
      return; 
    } 
    if (blogUrl.includes("blog.naver.com")) { 
      profileCell.setValue("🔄 프로필 로딩 중..."); 
      SpreadsheetApp.flush(); 
      const profileImg = getBloggerMainProfile(blogUrl); 
      if (profileImg) { 
        profileCell.setValue(profileImg); 
      } else { 
        profileCell.setValue("Main"); 
      } 
    } 
  } 
}

/******************************************************************************************************************
 * [내부 로직] 네이버 블로그 프로필 이미지 크롤링
 *****************************************************************************************************************/
function getBloggerMainProfile(url) { 
  let blogId = ""; 
  if (url.includes("blog.naver.com")) { 
    let parts = url.split("?")[0].replace("https://", "").replace("http://", "").split("/"); 
    if (parts[0].includes("blog.naver.com") && parts[1]) { blogId = parts[1]; } 
    else { let idMatch = url.match(/blogId=([^&]+)/); if (idMatch) blogId = idMatch[1]; } 
  } 
  if (!blogId) return ""; 
  let targetUrl = "https://m.blog.naver.com/" + blogId; 
  try { 
    let response = UrlFetchApp.fetch(targetUrl, { muteHttpExceptions: true, headers: { "User-Agent": "Mozilla/5.0 (Linux; Android 10; K)" } }); 
    let html = response.getContentText(); 
    let profileImgSrc = ""; 
    let profileMatch = html.match(/<meta\s+property=["']og:image["']\s+content=["']([^"']+)["']/i) || html.match(/<meta\s+property=["']naverblog:profile_image["']\s+content=["']([^"']+)["']/i) || html.match(/class=["']th_thumb["'][^>]*src=["']([^"']+)["']/i); 
    if (profileMatch) { 
      profileImgSrc = profileMatch[1]; 
      if (profileImgSrc.includes("buddy_profile") || profileImgSrc.includes("static.naver")) { return "Main"; } 
      if (profileImgSrc.includes("type=")) { profileImgSrc = profileImgSrc.replace(/type=[a-zA-Z0-9_]+/, "type=w3840"); } 
      else { profileImgSrc += (profileImgSrc.includes("?") ? "&" : "?") + "type=w3840"; } 
    } 
    return profileImgSrc || "Main"; 
  } catch (e) { return "Main"; } 
}

/******************************************************************************************************************
 * [공통 유틸리티] 날짜 포맷 변환 함수
 *****************************************************************************************************************/
function formatToShortDate(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "M/d");
  return val.toString().trim();
}

/****************************************************************************************************************
 * [공통 유틸리티] 유튜브 ID 추출 함수
 *****************************************************************************************************************/
function extractYouTubeId(url) {
  const regExp = /^.*((youtu.be\/)|(v\/)|(\/u\/\w\/)|(embed\/)|(watch\??v?=\??))([^#\&\?]*).*/;
  const match = url.match(regExp);
  return (match && match[7].length == 11) ? match[7] : false;
}
