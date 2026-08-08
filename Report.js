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
  const syncResult = syncBloggerDataOptimized(startRow, targetClient);
  const endTime = new Date().getTime();
  const duration = ((endTime - startTime) / 1000).toFixed(1);

  // 동기화 결과 상세 HTML 팝업 출력
  showSyncResultModal(syncResult, startRow, duration);
}

/*********************************************************************************
 * [결과 안내 HTML 모달]
 *********************************************************************************/
function showSyncResultModal(resultList, startRow, duration) {
  let totalCount = 0;
  let listHtml = "";

  if (!Array.isArray(resultList) || resultList.length === 0) {
    listHtml = `<div style="color: #64748b; padding: 20px 0; text-align: center;">새롭게 추가되거나 변경된 데이터가 없습니다. (모두 이미 완료됨)</div>`;
  } else {
    resultList.forEach(item => {
      totalCount += item.count;
      listHtml += `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px solid #f1f5f9;">
          <div>
            <strong style="color: #0f172a; font-size: 14px;">📍 ${item.clientName}</strong>
            <span style="background: #e2e8f0; color: #334155; font-size: 11px; padding: 2px 6px; border-radius: 4px; margin-left: 6px;">${item.count}건 처리</span>
          </div>
          <a href="${item.url}" target="_blank" style="background: #2563eb; color: white; padding: 5px 10px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: bold;">🔗 레포트 열기</a>
        </div>
      `;
    });
  }

  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: sans-serif; margin: 0; padding: 16px; color: #0f172a; }
        .header { margin-bottom: 16px; }
        .summary { background: #f8fafc; padding: 12px; border-radius: 8px; font-size: 13px; color: #475569; margin-bottom: 16px; border: 1px solid #e2e8f0; }
        .list-container { max-height: 280px; overflow-y: auto; }
        .btn-close { width: 100%; margin-top: 16px; padding: 10px; background: #0f172a; color: white; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
      </style>
    </head>
    <body>
      <div class="header">
        <h3 style="margin: 0 0 4px 0;">🎉 동기화 완료 리포트</h3>
        <p style="margin: 0; font-size: 12px; color: #64748b;">${startRow}행부터 동기화가 성공적으로 완료되었습니다.</p>
      </div>
      <div class="summary">
        ⏱️ 소요 시간: <strong>${duration}초</strong> | ✅ 신규/업데이트: <strong>${totalCount}건</strong>
      </div>
      <div class="list-container">
        ${listHtml}
      </div>
      <button class="btn-close" onclick="google.script.host.close()">창 닫기</button>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(420)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '업체별 동기화 결과');
}

/*********************************************************************************
 * [스마트 동적 매핑]
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

  for (let i = 1; i < data.length; i++) {
    const rawJpNameWithSama = data[i][2] ? data[i][2].toString().trim() : ""; // C열
    const rawJpName = data[i][3] ? data[i][3].toString().trim() : "";         // D열
    const rawKrName = data[i][4] ? data[i][4].toString().trim() : "";         // E열
    const status = data[i][6] ? data[i][6].toString().trim() : "진행중";    // G열
    let rawSheetId = data[i][11] ? data[i][11].toString().trim() : "";    // L열

    if (!rawSheetId || status === "계약종료" || status === "중단" || status === "OFF") {
      continue;
    }

    if (rawSheetId.includes("/d/")) {
      const match = rawSheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) rawSheetId = match[1];
    }

    const baseNames = [rawJpNameWithSama, rawJpName, rawKrName].filter(Boolean);

    baseNames.forEach(name => {
      clientMap[name] = rawSheetId;
      clientMap[name.replace(/\s+/g, '')] = rawSheetId;
      
      let clean = name.replace("クライアント名+様", "").replace(/様$/g, "").trim();
      if (clean) {
        clientMap[clean] = rawSheetId;
        clientMap[clean + "様"] = rawSheetId;
        clientMap[clean + " 様"] = rawSheetId;
      }
    });
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
  const cellsToHighlight = [];

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
        if (!updateBundles[sheetId]) {
          updateBundles[sheetId] = {
            clientName: storeName,
            rows: []
          };
        }
        
        updateBundles[sheetId].rows.push([
          scheduleDate, currentProfile, bloggerName, storeName,
          formatToShortDate(row[dateCol]), articleUrl, currentReports[j] || ""
        ]);

        cellsToHighlight.push({ row: i + 1, col: urlCol + 1 });
      }
    }
  }

  const resultList = [];

  for (const sheetId in updateBundles) {
    const bundle = updateBundles[sheetId];
    const count = processBatchUpdate(sheetId, bundle.rows);
    if (count > 0) {
      resultList.push({
        clientName: bundle.clientName,
        count: count,
        url: `https://docs.google.com/spreadsheets/d/${sheetId}`
      });
    }
  }

  if (cellsToHighlight.length > 0) {
    cellsToHighlight.forEach(pos => {
      masterSheet.getRange(pos.row, pos.col).setBackground("#e6f4ea");
    });
  }

  return resultList;
}

/*********************************************************************************
 * URL 정제 헬퍼 함수 (중복 비교 정밀도 향상)
 *********************************************************************************/
function cleanUrlForComparison(url) {
  if (!url) return "";
  return url.toString().trim()
            .toLowerCase()
            .replace(/^https?:\/\//i, '')
            .replace(/^m\.blog\.naver\.com/i, 'blog.naver.com')
            .replace(/\/$/, '');
}

/*********************************************************************************
 * 특정 시트에 모인 데이터를 한꺼번에 업데이트 (날짜 타입 비교 오류 수정 버전)
 *********************************************************************************/
function processBatchUpdate(sheetId, newRows) {
  try {
    const targetSs = SpreadsheetApp.openById(sheetId);
    const sheet = targetSs.getSheetByName(new Date().getFullYear() + "年") || targetSs.getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    // 💡 getValues() 대신 getDisplayValues()를 사용하여 화면에 보이는 텍스트로 정밀 비교
    const existingData = lastRow >= 2 ? sheet.getRange(1, 1, lastRow, 7).getDisplayValues() : [];
    
    let updatedCount = 0;
    const toAppend = [];

    // 기존 시트에 이미 작성된 포스팅 URL 맵 생성
    const existingUrlMap = {};
    for (let i = 0; i < existingData.length; i++) {
      const existingUrlKey = cleanUrlForComparison(existingData[i][5]);
      if (existingUrlKey) {
        existingUrlMap[existingUrlKey] = {
          rowIndex: i + 1,
          rowValues: existingData[i]
        };
      }
    }

    newRows.forEach(rowData => {
      const targetUrlKey = cleanUrlForComparison(rowData[5]); // articleUrl
      
      if (targetUrlKey && existingUrlMap[targetUrlKey]) {
        // 1. 이미 동기화된 URL이 존재할 경우 -> 진짜 내용 변경이 있을 때만 업데이트
        const match = existingUrlMap[targetUrlKey];
        const oldRow = match.rowValues;
        
        const isChanged = 
          formatToShortDate(oldRow[0]) !== formatToShortDate(rowData[0]) || // 일정 날짜
          String(oldRow[2]).trim() !== String(rowData[2]).trim() ||         // 작성자
          formatToShortDate(oldRow[4]) !== formatToShortDate(rowData[4]) || // 작성일
          String(oldRow[6]).trim() !== String(rowData[6]).trim();           // 보고서 URL
          
        if (isChanged) {
          sheet.getRange(match.rowIndex, 1, 1, 7).setValues([rowData]);
          updatedCount++;
        }
        // 변경 사항이 없으면 아무 작업도 하지 않고 스킵 (0건 처리)
      } else {
        // 2. 레포트 시트에 없는 신규 URL만 추가
        toAppend.push(rowData);
      }
    });

    if (toAppend.length > 0) {
      sheet.getRange(lastRow + 1, 1, toAppend.length, 7).setValues(toAppend);
      updatedCount += toAppend.length;
    }

    const newLastRow = sheet.getLastRow();
    if (newLastRow >= 2 && updatedCount > 0) {
      sheet.getRange(2, 1, newLastRow - 1, 7).sort([
        {column: 1, ascending: true}, 
        {column: 4, ascending: true}, 
        {column: 3, ascending: true}  
      ]);
    } 

    return updatedCount;
  } catch (e) {
    console.error("ID 확인 불가 시트 스킵: " + sheetId + " | 오류: " + e.message);
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
    { url: 9,  title: 11, jp: 12 },
    { url: 23, title: 25, jp: 26 }
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
