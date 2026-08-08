/*********************************************************************************
 * [보고서 자동화] 메뉴 실행 및 메인 제어
 *********************************************************************************/
function runSyncWithPrompt() {
  const ui = SpreadsheetApp.getUi();
  
  const rowResponse = ui.prompt('1. 동기화 시작 행 입력', '시작 행 번호 (예: 67)', ui.ButtonSet.OK_CANCEL);
  if (rowResponse.getSelectedButton() !== ui.Button.OK) return;
  const startRow = parseInt(rowResponse.getResponseText());

  const clientResponse = ui.prompt('2. 업데이트할 클라이언트명 입력', '특정 업체만 하려면 입력, 전체는 비워두세요.', ui.ButtonSet.OK_CANCEL);
  if (clientResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetClient = clientResponse.getResponseText().trim();

  if (isNaN(startRow) || startRow < 1) {
    ui.alert('올바른 행 번호를 입력해주세요.');
    return;
  }
  
  const startTime = new Date().getTime();
  const syncedCount = syncBloggerDataOptimized(startRow, targetClient);
  const endTime = new Date().getTime();
  
  ui.alert(`${startRow}행부터 완료!\n✅ 처리 건수: ${syncedCount}건\n⏱️ 소요 시간: ${(endTime - startTime)/1000}초`);
}

/*********************************************************************************
 * [동적 매핑] 클라이언트 Master Data 시트에서 업체별 시트 ID Map 동적 생성
 *********************************************************************************/
function getDynamicClientMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!masterSheet) {
    console.error("❌ '클라이언트 Master Data' 시트를 찾을 수 없습니다.");
    return {};
  }

  const data = masterSheet.getDataRange().getValues();
  const clientMap = {};

  // 헤더(1행) 제외하고 2행부터 순회
  for (let i = 1; i < data.length; i++) {
    const jpName = data[i][2] ? data[i][2].toString().trim() : ""; // C열: 일문명
    const krName = data[i][4] ? data[i][4].toString().trim() : ""; // E열: 한글명
    let rawSheetId = data[i][5] ? data[i][5].toString().trim() : ""; // F열: 시트 ID 또는 URL

    if (!rawSheetId) continue;

    // URL 형식일 경우 정규표현식으로 시트 ID만 정밀 추출
    if (rawSheetId.includes("/d/")) {
      const match = rawSheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) rawSheetId = match[1];
    }

    // 일문명과 한글명 모두 Map 키값으로 등록 (둘 중 어떤 이름이 들어와도 인식 가능)
    if (jpName) clientMap[jpName] = rawSheetId;
    if (krName) clientMap[krName] = rawSheetId;
  }

  return clientMap;
}

/*********************************************************************************
 * 최적화된 메인 동기화 로직
 *********************************************************************************/
function syncBloggerDataOptimized(startRow, targetClient) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_SCHEDULE); 
  const lastRow = masterSheet.getLastRow();
  const data = masterSheet.getRange(1, 1, lastRow, masterSheet.getLastColumn()).getValues();
  
  // 🚀 시트 기반 동적 클라이언트 매핑 적용
  const clientMap = getDynamicClientMap();
  
  const updateBundles = {}; 
  const cleanTarget = targetClient ? targetClient.replace(/\s+/g, '') : null;

  function preScanReports(hIdx) {
    let reports = Array(10).fill("");
    for (let k = hIdx; k < Math.min(hIdx + 30, lastRow); k++) {
      if (data[k][2] && data[k][2].toString().trim() === "報告書作成") { 
        for (let j = 0; j < 10; j++) {
          let reportCol = 8 + (j * 3); 
          let rUrl = data[k][reportCol] ? data[k][reportCol].toString().trim() : "";
          
          if (rUrl.includes("http")) {
            reports[j] = rUrl;
          } else if (rUrl.includes("file/d/")) {
            reports[j] = "https://drive.google.com/" + rUrl;
          }
        }
        break;
      }
    }
    return reports;
  }

  let headerIdx = -1;
  let scheduleDate = "";
  let currentProfile = "";
  let currentClients = [];
  let currentReports = [];

  for (let i = 0; i < lastRow; i++) {
    const row = data[i];
    if (row[4] && row[4].toString().includes("Schedule")) { 
      headerIdx = i;
      scheduleDate = formatToShortDate(row[2]); 
      currentProfile = row[3] ? row[3].toString().trim() : ""; 
      currentClients = [];
      for (let j = 0; j < 10; j++) {
        currentClients.push(row[9 + (j * 3)] ? row[9 + (j * 3)].toString().trim() : ""); 
      }
      currentReports = preScanReports(headerIdx);
      continue;
    }

    if (i < startRow - 1 || headerIdx === -1) continue;

    const bloggerName = row[2] ? row[2].toString().trim() : ""; 
    if (!bloggerName || bloggerName === "報告書作成" || bloggerName.includes("報告") || bloggerName.includes("BLOG") || bloggerName.includes("REPORT")) continue;

    for (let j = 0; j < 10; j++) {
      const storeName = currentClients[j];
      if (!storeName || !clientMap[storeName]) continue;
      if (cleanTarget && storeName.replace(/\s+/g, '') !== cleanTarget) continue;

      const dateCol = 8 + (j * 3); 
      const urlCol = 9 + (j * 3);  
      
      const articleUrl = row[urlCol] ? row[urlCol].toString().trim() : "";
      
      if (articleUrl.includes("http")) {
        const sheetId = clientMap[storeName];
        if (!updateBundles[sheetId]) updateBundles[sheetId] = [];
        
        updateBundles[sheetId].push([
          scheduleDate, currentProfile, bloggerName, storeName,
          formatToShortDate(row[dateCol]), articleUrl, currentReports[j] || ""
        ]);
      }
    }
  }

  let totalCount = 0;
  for (const sheetId in updateBundles) {
    totalCount += processBatchUpdate(sheetId, updateBundles[sheetId]);
  }

  return totalCount;
}

/*********************************************************************************
 * 특정 시트에 모인 데이터를 한꺼번에 업데이트
 *********************************************************************************/
function processBatchUpdate(sheetId, newRows) {
  try {
    const targetSs = SpreadsheetApp.openById(sheetId);
    const sheet = targetSs.getSheetByName(new Date().getFullYear() + "年") || targetSs.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const existingData = lastRow >= 2 ? sheet.getRange(1, 1, lastRow, 7).getValues() : [];
    
    let updatedCount = 0;
    const toAppend = [];

    newRows.forEach(rowData => {
      let isFound = false;
      for (let i = 0; i < existingData.length; i++) {
        if (formatToShortDate(existingData[i][0]) === rowData[0] &&
            String(existingData[i][2]).trim() === rowData[2] &&
            String(existingData[i][3]).trim() === rowData[3] &&
            String(existingData[i][5]).trim() === rowData[5]) {
          
          isFound = true;
          if (String(existingData[i][6]).trim() !== String(rowData[6]).trim() || 
              String(existingData[i][4]).trim() !== String(rowData[4]).trim()) {
            sheet.getRange(i + 1, 1, 1, 7).setValues([rowData]);
            updatedCount++;
          }
          break;
        }
      }
      if (!isFound) toAppend.push(rowData);
    });

    if (toAppend.length > 0) {
      sheet.getRange(lastRow + 1, 1, toAppend.length, 7).setValues(toAppend);
      updatedCount += toAppend.length;
    }

    const newLastRow = sheet.getLastRow();
    if (newLastRow >= 2) {
      sheet.getRange(2, 1, newLastRow - 1, 7).sort([
        {column: 1, ascending: true}, 
        {column: 4, ascending: true}, 
        {column: 3, ascending: true}  
      ]);
    } 

    return updatedCount;
  } catch (e) {
    console.error("ID 확인 불가 시트 스킵: " + sheetId);
    return 0;
  }
}

/*********************************************************************************
 * [관리자 기능] 블로그 타이틀 자동 추출
 *********************************************************************************/
function fetchAndTranslateBlogTitles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.REPORT);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 보고서 시트를 찾을 수 없습니다.");
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  
  if (startRow < 12) {
    SpreadsheetApp.getUi().alert("⚠️ 실제 데이터가 있는 12번째 행부터 선택해 주세요.");
    return;
  }
  
  let successCount = 0;
  
  const targetColumns = [
    { url: 9,  title: 11, jp: 12 }, // 1번 블로그 (I, K, L열)
    { url: 23, title: 25, jp: 26 } // 2번 블로그 (W, Y, Z열)
  ];
  
  for (let i = 0; i < numRows; i++) {
    let currentRow = startRow + i;
    
    targetColumns.forEach(cols => {
      let url = sheet.getRange(currentRow, cols.url).getValue().toString().trim();
      if (url.includes("blog.naver.com")) {
        let title = getNaverBlogRealTitle(url); 
        if (title) {
          sheet.getRange(currentRow, cols.title).setValue(title); 
          sheet.getRange(currentRow, cols.jp).setValue("");    
        }
        successCount++;
      }
    });
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(`🎉 총 ${successCount}건의 한국어 타이틀 수집이 완료되었습니다!`);
}

/*********************************************************************************
 * 네이버 블로그 제목 추출 내부 함수
 *********************************************************************************/
function getNaverBlogRealTitle(url) {
  let blogId = ""; let logNo = "";
  if (url.includes("PostView.naver") || url.includes("PostView.nhn")) {
    let idMatch = url.match(/blogId=([^&]+)/); let logMatch = url.match(/logNo=([^&]+)/);
    if (idMatch) blogId = idMatch[1]; if (logMatch) logNo = logMatch[1];
  } else {
    let parts = url.split("?")[0].replace("https://", "").replace("http://", "").split("/");
    if (parts.length >= 3) { blogId = parts[1]; logNo = parts[2]; }
  }
  if (!blogId || !logNo) return "";

  let targetUrl = "https://blog.naver.com/PostView.naver?blogId=" + blogId + "&logNo=" + logNo;
  try {
    let response = UrlFetchApp.fetch(targetUrl, { muteHttpExceptions: true });
    let html = response.getContentText();
    let titleMatch = html.match(/<meta\s+property=["']og:title["']\s+content=["']([^"']+)["']/i) || html.match(/<title>([^<]+)<\/title>/i);
                     
    if (titleMatch) {
      let cleanTitle = titleMatch[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#39;/g, "'").replace(/&apos;/g, "'").replace(/&times;/g, "×"); 
      return cleanTitle.replace(" : 네이버 블로그", "").trim();
    }
    return "";
  } catch (e) { return ""; }
}
