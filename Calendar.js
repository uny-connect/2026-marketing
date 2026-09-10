/*********************************************************************************
 * [일본 취재 등록봇] 2026スケジュール 자동화 시스템 (마지막 날 제외 버전)
 *********************************************************************************/
function checkAndRegisterSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.JAPAN_SCHEDULE);
  const calendarName = CONFIG.CALENDARS.JAPAN_TRIP;
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  
  if (!sheet || calendars.length === 0) {
    showAlertSafely("시트명이나 캘린더명을 확인해주세요.");
    return;
  }

  const calendar = calendars[0];
  const today = new Date();
  today.setHours(0, 0, 0, 0); 
  const year = 2026;

  // C, I, O, U, AA, AG, AM 열 (시트 기준 3, 9, 15, 21, 27, 33, 39열)
  const startCols = [3, 9, 15, 21, 27, 33, 39]; 

  const sheetData = sheet.getDataRange().getDisplayValues();
  let registerCount = 0;

  // M/D 또는 M/D-D 형식 패턴 감지
  const dateRegex = /^(\d{1,2})\/(\d{1,2})(-\d{1,2})?$/;

  startCols.forEach(baseCol => {
    const colIdx = baseCol - 1;

    for (let rowIdx = 0; rowIdx < sheetData.length; rowIdx++) {
      if (colIdx >= sheetData[rowIdx].length) continue;

      const dateRaw = sheetData[rowIdx][colIdx].trim();
      if (!dateRegex.test(dateRaw)) continue;

      try {
        const parts = dateRaw.split('/');
        const month = parseInt(parts[0], 10);
        const dayParts = parts[1].split('-');
        const startDay = parseInt(dayParts[0], 10);
        const endDay = parseInt(dayParts[1] || dayParts[0], 10);

        const startDate = new Date(year, month - 1, startDay);
        
        // 💡 마지막 날 제외 처리: 9/9-12라면 11일까지만 등록 (시작일과 종료일이 같으면 당일 등록)
        let actualEndDay = endDay;
        if (endDay > startDay) {
          actualEndDay = endDay - 1;
        }
        const calendarEndDate = new Date(year, month - 1, actualEndDay);

        const diffDays = Math.ceil((startDate - today) / (1000 * 60 * 60 * 24));

        // 오늘 기준 -1일(어제)부터 30일 이내의 일정만 등록
        if (diffDays <= 30 && diffDays >= -1) {
          let staffList = [];
          let dailySchedules = {}; 
          let lastDay = startDay + "日"; 
          dailySchedules[lastDay] = [];

          // 날짜 행 바로 아래부터 최대 25행 내용 추출
          const contentData = [];
          for (let r = 1; r <= 25; r++) {
            const currR = rowIdx + r;
            if (currR < sheetData.length) {
              const rowVals = [];
              for (let c = 0; c < 4; c++) {
                const currC = colIdx + c;
                rowVals.push(currC < sheetData[currR].length ? sheetData[currR][currC] : "");
              }
              contentData.push(rowVals);
            }
          }
          
          contentData.forEach((row) => {
            let col1 = row[0].toString().trim(); 
            let col2 = row[1].toString().trim(); 
            let col3 = row[2].toString().trim(); 
            let col4 = row[3].toString().trim(); 

            let onlyNum = col1.replace(/[^0-9]/g, "");
            if (onlyNum !== "" && col1.length <= 3) {
              let dayNum = parseInt(onlyNum, 10);
              if (dayNum >= 1 && dayNum <= 31) {
                lastDay = dayNum + "日";
                if (!dailySchedules[lastDay]) dailySchedules[lastDay] = [];
              }
            }

            const skipWords = ["名前", "URL", "空港", "備考", "No", "店名", "순번"];
            if (skipWords.includes(col1) || col1.includes('/')) return;

            // 담당자/스태프 및 비행기 착륙 정보 추출
            if (col1 !== "" && isNaN(parseInt(col1.replace(/[^0-9]/g, ""), 10))) {
              let staffName = col1;
              const arrivalKeywords = ["着陸", "🛬", "🛫", "離陸", "착륙", "이륙"];
              const hasArrivalInfo = arrivalKeywords.some(kw => col2.includes(kw));
              
              if (col2 && hasArrivalInfo) {
                staffName += ` ${col2}`;
              }
              
              if (staffList.indexOf(staffName) === -1) staffList.push(staffName);
            }

            // 매장 일정 수집
            if (col4 !== "" && col4 !== "備考" && col4 !== "店名") {
              let circleNo = toCircleNumberExtended(col3); 
              let timeStr = "";
              
              if (col2.includes("宿泊") || col2.includes("숙박")) {
                timeStr = "🏨 ";
              } else if (col2 && !col2.includes("URL") && !col2.includes("着陸") && !col2.includes("🛬") && !col2.includes("🛫")) {
                timeStr = col2 + " ";
              }
              
              let entry = `${circleNo}${timeStr}${col4}`.trim();
              if (entry) dailySchedules[lastDay].push(entry);
            }
          });

          let scheduleParts = [];
          let sortedDays = Object.keys(dailySchedules).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));
          
          sortedDays.forEach(dayKey => {
            if (dailySchedules[dayKey].length > 0) {
              scheduleParts.push(`🗓️${dayKey}: ${dailySchedules[dayKey].join(', ')}`);
            }
          });

          const staffInfo = staffList.length > 0 ? ` (👤${staffList.join(',')})` : "";
          const spacer = "                    "; 
          const finalTitle = `[韓国ブロガー招聘] ${dateRaw}${staffInfo}${spacer}${scheduleParts.join(spacer)}`;

          // 💡 calendarEndDate(마지막 전날)까지만 루프 생성
          for (let d = new Date(startDate); d <= calendarEndDate; d.setDate(d.getDate() + 1)) {
            const currentDayStart = new Date(d);
            currentDayStart.setHours(7, 0, 0); 
            const currentDayEnd = new Date(d);
            currentDayEnd.setHours(22, 0, 0); 

            const existingEvents = calendar.getEvents(currentDayStart, currentDayEnd, {search: `[韓国ブロガー招聘] ${dateRaw}`});
            if (existingEvents.length === 0) {
              calendar.createEvent(finalTitle, currentDayStart, currentDayEnd, { description: "" });
              registerCount++;
            }
          }
        }
      } catch(e) { 
        console.log(`[오류] 행 ${rowIdx+1}: ${e.message}`); 
      }
    }
  });

  showAlertSafely(`일정 동기화가 완료되었습니다. (신규 등록: ${registerCount}건)`);
}

/**
 * 트리거 실행(UI 없음)과 수동 실행(UI 있음) 모두 안전하게 처리하는 알림 함수
 */
function showAlertSafely(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    console.log("[트리거 자동 실행 로그] " + msg);
  }
}

/**
 * 원형 숫자 변환 내부 함수
 */
function toCircleNumberExtended(num) {
  if (!num) return "";
  let raw = num.toString().trim();
  if (raw.includes('.') || raw.includes('-')) return `[${raw}] `;
  let n = parseInt(raw.replace(/[^0-9]/g, ""), 10);
  if (isNaN(n) || n <= 0) return ""; 
  if (n >= 1 && n <= 10) return String.fromCharCode(0x2460 + n - 1) + " ";
  return `[${n}] `; 
}

/*********************************************************************************
 * [취재 마감 알람 등록봇] 캘린더봇 시트 연동 시스템
 *********************************************************************************/
function createCoverageSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR_BOT);

  if (!sheet) {
    showAlertSafely("❌ '캘린더봇' 시트를 찾을 수 없습니다.");
    return;
  }

  const calendarName = CONFIG.CALENDARS.COVERAGE; 
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length === 0) {
    showAlertSafely(`❌ '${calendarName}' 캘린더를 찾을 수 없습니다.`);
    return;
  }
  const calendar = calendars[0];
  const data = sheet.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const title = data[i][0];
    const startDate = new Date(data[i][1]);
    const endDate = new Date(data[i][2]);
    const status = data[i][3];

    if (status === "완료" || !title || isNaN(startDate.getTime())) continue;

    try {
      let alarmDayBefore = new Date(startDate);
      alarmDayBefore.setDate(startDate.getDate() - 1);
      createAtNineAM(calendar, `[알람] ${title} 하루 전`, alarmDayBefore, alarmDayBefore);

      createAtNineAM(calendar, title, startDate, endDate);

      let deadlineOneWeek = new Date(endDate);
      deadlineOneWeek.setDate(endDate.getDate() + 8);
      createAtNineAM(calendar, `${title} 마감 1주일전`, deadlineOneWeek, deadlineOneWeek);

      let deadlineFinal = new Date(deadlineOneWeek);
      deadlineFinal.setDate(deadlineOneWeek.getDate() + 8);
      createAtNineAM(calendar, `${title} 마감 당일`, deadlineFinal, deadlineFinal);

      sheet.getRange(i + 1, 4).setValue("완료");
      count++;
    } catch (e) {
      console.log(`${i+1}행 등록 중 오류: ${e.message}`);
    }
  }
  showAlertSafely(`✅ 총 ${count}건의 취재가 구글캘린더에 등록되었습니다.`);
}

/**
 * 당일 오전 09:00 정시 알람 일정 생성 헬퍼 함수
 */
function createAtNineAM(calendar, title, startD, endD) {
  const sDate = new Date(startD);
  sDate.setHours(9, 0, 0, 0);
  
  const eDate = new Date(endD);
  eDate.setHours(10, 0, 0, 0);

  const event = calendar.createEvent(title, sDate, eDate);
  event.removeAllReminders(); 
  event.addPopupReminder(0);  
}

/**
 * 매일 새벽 자동 실행 트리거 등록 함수
 */
function createJapanTripTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndRegisterSchedules') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkAndRegisterSchedules')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  showAlertSafely('✅ 일본 취재 등록봇의 매일 새벽 자동 실행 트리거가 정상 설정되었습니다.');
}
