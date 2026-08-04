/*********************************************************************************
 * [웹앱 기능] 크리에이터 접속 시 화면 출력 (doGet)
 *********************************************************************************/
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('크리에이터 콘텐츠 제출')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/**
 * [웹앱 기능] 크리에이터 코드 검증 및 매장 매칭
 */
function verifyCode(inputCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // CONFIG 설정값 반영
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_SCHEDULE); 
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA); 
  if (!sheet) return { success: false, message: "오류: 스케줄 시트를 찾을 수 없습니다." };
  
  const nameMap = {};
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    for (let m = 0; m < masterData.length; m++) {
      let jpName = masterData[m][2] ? masterData[m][2].toString().trim() : ""; 
      let krName = masterData[m][4] ? masterData[m][4].toString().trim() : ""; 
      if (jpName && krName) nameMap[jpName] = krName;
    }
  }

  const data = sheet.getDataRange().getValues();
  const lastCol = sheet.getLastColumn();
  const targetCode = inputCode.toString().trim();
  
  for (let i = 0; i < data.length; i++) {
    let sheetCode = data[i][1] ? data[i][1].toString().trim() : "";
    if (sheetCode === targetCode) {
      let creatorName = data[i][2] ? data[i][2].toString().trim() : "크리에이터"; 
      let headerRowIdx = -1;
      for (let r = i; r >= 0; r--) {
        if (data[r] && data[r][8] && data[r][8].toString().trim() === "投稿日") {
          headerRowIdx = r - 1; break;
        }
      }

      let clients = [];
      if (headerRowIdx >= 0) {
        for (let col = 8; col < lastCol; col += 3) {
          let rawClientName = (data[headerRowIdx] && data[headerRowIdx].length > col + 1) ? data[headerRowIdx][col + 1].toString().trim() : "";
          if (rawClientName) {
            let displayClientName = nameMap[rawClientName] || rawClientName;
            let urlCellVal = (data[i] && data[i].length > col + 1) ? data[i][col + 1].toString().trim() : "";
            let isExcluded = (urlCellVal === "-");
            let existingUrl = (!isExcluded && urlCellVal.includes("http")) ? urlCellVal : "";

            clients.push({
              colIndex: col + 1, 
              name: displayClientName,
              excluded: isExcluded,
              existingUrl: existingUrl
            });
          }
        }
      }
      return { success: true, row: i + 1, creatorName: creatorName, clients: clients };
    }
  }
  return { success: false, message: "일치하는 코드가 없습니다. 다시 확인해주세요." };
}

/**************************************************************************************************************
 * [웹앱 기능] 크리에이터 링크 제출 처리
 **************************************************************************************************************/
function saveSingleUrl(row, colIndex, url) {
  // CONFIG 설정값 반영
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MASTER_SCHEDULE);
  if (!sheet) throw new Error("시트를 찾을 수 없습니다.");

  const urlCol = colIndex + 1; 
  const dateCol = urlCol - 1;       
  const imgCol = urlCol + 1;        

  sheet.getRange(row, urlCol).setValue(url);

  let pubDate = "";
  let imgUrl = "";

  try {
    if (url.includes("blog.naver.com")) {
      const blogData = getNaverBlogData(url);
      pubDate = blogData.date; 
      imgUrl = blogData.img;
    } else if (url.includes("youtube.com") || url.includes("youtu.be")) {
      const videoId = extractYouTubeId(url); // 외부 정의 함수 호출유지
      if (videoId) imgUrl = "https://img.youtube.com/vi/" + videoId + "/0.jpg";
    } else if (url.includes("instagram.com")) {
      imgUrl = "Main"; 
    }
  } catch (e) {
    console.log("크롤링 실패: " + e.message);
  }

  if (pubDate) {
    sheet.getRange(row, dateCol).setValue(pubDate);
  }

  if (!imgUrl) imgUrl = "Main";
  sheet.getRange(row, imgCol).setValue(imgUrl);

  return { success: true, url: url };
}

/****************************************************************************************************************
 * [내부 기능] 네이버 블로그 상세 데이터 추출 (시간/분 정밀 역산 반영 버전)
 ***************************************************************************************************************/
function getNaverBlogData(url) {
  let blogId = ""; let logNo = "";
  if (url.includes("PostView.naver") || url.includes("PostView.nhn")) {
    let idMatch = url.match(/blogId=([^&]+)/); let logMatch = url.match(/logNo=([^&]+)/);
    if (idMatch) blogId = idMatch[1]; if (logMatch) logNo = logMatch[1];
  } else {
    let parts = url.split("?")[0].replace("https://", "").replace("http://", "").split("/");
    if (parts.length >= 3) { blogId = parts[1]; logNo = parts[2]; }
  }
  if (!blogId || !logNo) return { date: "", img: "Main", profileImg: "" };
  let targetUrl = "https://blog.naver.com/PostView.naver?blogId=" + blogId + "&logNo=" + logNo;
  try {
    let response = UrlFetchApp.fetch(targetUrl, { muteHttpExceptions: true });
    let html = response.getContentText();
    let imgSrc = "";
    const postfileMatch = html.match(/(https?:\/\/postfiles\.pstatic\.net\/[^"']+)/);
    if (postfileMatch) { imgSrc = postfileMatch[1]; if (!imgSrc.includes("type=")) { imgSrc += (imgSrc.includes("?") ? "&" : "?") + "type=w966"; } } 
    else { let imgMatch = html.match(/<meta property="og:image" content="([^"]+)"/i); if (imgMatch) imgSrc = imgMatch[1]; }

    let profileImgSrc = "";
    let mobileUrl = "https://m.blog.naver.com/" + blogId;
    let mResponse = UrlFetchApp.fetch(mobileUrl, { muteHttpExceptions: true, headers: { "User-Agent": "Mozilla/5.0 (Linux; Android 10; K)" } });
    let mHtml = mResponse.getContentText();
    let profileMatch = mHtml.match(/<meta\s+property=["']og:image["']\s+content=["']([^"']+)["']/i) || mHtml.match(/<meta\s+property=["']naverblog:profile_image["']\s+content=["']([^"']+)["']/i);
    if (profileMatch) {
      profileImgSrc = profileMatch[1];
      if (!profileImgSrc.includes("buddy_profile") && !profileImgSrc.includes("static.naver")) {
        profileImgSrc = profileImgSrc.includes("type=") ? profileImgSrc.replace(/type=[a-zA-Z0-9_]+/, "type=w3840") : profileImgSrc + "?type=w3840";
      } else { profileImgSrc = "Main"; }
    }

    let pubDate = ""; 
    let dateMatch = html.match(/property="article:published_time"\s*content="([^"]+)"/i);
    
    if (dateMatch && dateMatch[1] && dateMatch[1].includes('T')) { 
      pubDate = dateMatch[1].split('T')[0]; 
    } else {
      let altDateMatch = html.match(/se_publishDate[^>]*>([^<]+)<\/span>/i) || html.match(/blog-date[^>]*>([^<]+)<\/span>/i);
      if (altDateMatch) { 
        let rawDateStr = altDateMatch[1].trim(); 
        let now = new Date(); 
        
        if (rawDateStr.includes("방금") || rawDateStr.includes("초 전")) {
          pubDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (rawDateStr.includes("분 전")) {
          let minutesMatch = rawDateStr.match(/(\d+)\s*분\s*전/);
          if (minutesMatch) {
            let minutes = parseInt(minutesMatch[1]);
            let calculatedDate = new Date(now.getTime() - (minutes * 60 * 1000));
            pubDate = Utilities.formatDate(calculatedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            pubDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        } else if (rawDateStr.includes("시간 전")) {
          let hoursMatch = rawDateStr.match(/(\d+)\s*시간\s*전/);
          if (hoursMatch) {
            let hours = parseInt(hoursMatch[1]);
            let calculatedDate = new Date(now.getTime() - (hours * 60 * 60 * 1000));
            pubDate = Utilities.formatDate(calculatedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            pubDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        } else if (rawDateStr.includes("어제")) {
          let yesterday = new Date(now.getTime() - (24 * 60 * 60 * 1000));
          pubDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          let dStr = rawDateStr.replace(/\./g, '-').replace(/\s/g, ''); 
          let dMatch = dStr.match(/(\d{4}-\d{1,2}-\d{1,2})/); 
          if (dMatch) {
            let dateParts = dMatch[1].split('-');
            let formattedYear = dateParts[0];
            let formattedMonth = dateParts[1].length === 1 ? "0" + dateParts[1] : dateParts[1];
            let formattedDay = dateParts[2].length === 1 ? "0" + dateParts[2] : dateParts[2];
            pubDate = formattedYear + "-" + formattedMonth + "-" + formattedDay;
          }
        }
      }
    }
    return { date: pubDate, img: imgSrc, profileImg: profileImgSrc || "Main" };
  } catch (e) { 
    return { date: "", img: "Main", profileImg: "Main" }; 
  }
}
