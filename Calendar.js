/*********************************************************************************
 * [일본 취재 등록봇] 2026スケジュール 자동화 시스템
 *********************************************************************************/
function checkAndRegisterSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // CONFIG 설정값 반영
  const sheet = ss.getSheetByName(CONFIG.SHEETS.JAPAN_SCHEDULE);
  const calendarName = CONFIG.CALENDARS.JAPAN_TRIP;
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  
  if (!sheet || calendars.length === 0) {
    SpreadsheetApp.getUi().alert("시트명이나 캘린더명을 확인해주세요.");
    return;
  }

  const calendar = calendars[0];
  const today = new Date();
  today.setHours(0, 0, 0, 0); 
  const year = 2026;

  const startCols = [3, 9, 15, 21, 27, 33, 39]; 
  const startRows = [51, 53, 63, 74, 76, 86, 95, 97, 108, 116, 118, 137, 158, 179, 200, 221, 242, 263, 284]; 

  startRows.forEach(baseRow => {
    startCols.forEach(baseCol => {
      const dateRaw = sheet.getRange(baseRow, baseCol).getDisplayValue().trim();
      if (!dateRaw || !dateRaw.includes('/')) return;

      try {
        const parts = dateRaw.split('/');
        const month = parseInt(parts[0]);
        const dayParts = parts[1].split('-');
        const startDay = parseInt(dayParts[0]);
        const endDay = parseInt(dayParts[1] || dayParts[0]);

        const startDate = new Date(year, month - 1, startDay);
        const endDate = new Date(year, month - 1, endDay);
        const diffDays = Math.ceil((startDate - today) / (1000 * 60 * 60 * 24));

        if (diffDays <= 14 && diffDays >= -1) {
          let staffList = [];
          let dailySchedules = {}; 
          let lastDay = startDay + "日"; 
          dailySchedules[lastDay] = [];

          const contentData = sheet.getRange(baseRow, baseCol, 25, 4).getDisplayValues();
          
          contentData.forEach((row, index) => {
            let col1 = row[0].toString().trim(); 
            let col2 = row[1].toString().trim(); 
            let col3 = row[2].toString().trim(); 
            let col4 = row[3].toString().trim(); 

            let onlyNum = col1.replace(/[^0-9]/g, "");
            if (onlyNum !== "" && col1.length <= 3) {
              let dayNum = parseInt(onlyNum);
              if (dayNum >= 1 && dayNum <= 31) {
                lastDay = dayNum + "日";
                if (!dailySchedules[lastDay]) dailySchedules[lastDay] = [];
              }
            }

            const skipWords = ["名前", "URL", "空港", "備考", "No", "店名", "순번"];
            if (index === 0 || skipWords.includes(col1) || col1.includes('/')) return;

            if (col1 !== "" && isNaN(parseInt(col1.replace(/[^0-9]/g, "")))) {
              let staffName = col1;
              const arrivalKeywords = ["着陸", "🛬", "🛫", "離陸", "착륙", "이륙"];
              const hasArrivalInfo = arrivalKeywords.some(kw => col2.includes(kw));
              
              if (col2 && hasArrivalInfo) {
                staffName += ` ${col2}`;
              }
              
              if (staffList.indexOf(staffName) === -1) staffList.push(staffName);
            }

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
          let sortedDays = Object.keys(dailySchedules).sort((a, b) => parseInt(a) - parseInt(b));
          
          sortedDays.forEach(dayKey => {
            if (dailySchedules[dayKey].length > 0) {
              scheduleParts.push(`🗓️${dayKey}: ${dailySchedules[dayKey].join(', ')}`);
            }
          });

          const staffInfo = staffList.length > 0 ? ` (👤${staffList.join(',')})` : "";
          const spacer = "                    "; 
          const finalTitle = `[韓国ブロガー招聘] ${dateRaw}${staffInfo}${spacer}${scheduleParts.join(spacer)}`;

          for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
            const currentDayStart = new Date(d);
            currentDayStart.setHours(7, 0, 0); 
            const currentDayEnd = new Date(d);
            currentDayEnd.setHours(22, 0, 0); 

            const existingEvents = calendar.getEvents(currentDayStart, currentDayEnd, {search: `[韓国ブロガー招聘] ${dateRaw}`});
            if (existingEvents.length === 0) {
              calendar.createEvent(finalTitle, currentDayStart, currentDayEnd, { description: "" });
            }
          }
        }
      } catch(e) { console.log(e.message); }
    });
  });
  SpreadsheetApp.getUi().alert("착륙 정보가 포함된 일정이 등록되었습니다.");
}

/**
 * 원형 숫자 변환 내부 함수
 */
function toCircleNumberExtended(num) {
  if (!num) return "";
  let raw = num.toString().trim();
  if (raw.includes('.') || raw.includes('-')) return `[${raw}] `;
  let n = parseInt(raw.replace(/[^0-9]/g, ""));
  if (isNaN(n) || n <= 0) return ""; 
  if (n >= 1 && n <= 10) return String.fromCharCode(0x2460 + n - 1) + " ";
  return `[${n}] `; 
}

/*********************************************************************************
 * [취재 마감 알람 등록봇] 캘린더봇 시트 연동 시스템
 *********************************************************************************/
function createCoverageSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // CONFIG 설정값 반영
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR_BOT);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert("❌ '캘린더봇' 시트를 찾을 수 없습니다.");
    return;
  }

  const calendarName = CONFIG.CALENDARS.COVERAGE; 
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length === 0) {
    ui.alert(`❌ '${calendarName}' 캘린더를 찾을 수 없습니다.`);
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
      calendar.createAllDayEvent(`[알람] ${title} 하루 전`, alarmDayBefore);

      let calEndDate = new Date(endDate);
      calEndDate.setDate(endDate.getDate() + 1);
      calendar.createAllDayEvent(title, startDate, calEndDate);

      let deadlineOneWeek = new Date(endDate);
      deadlineOneWeek.setDate(endDate.getDate() + 8);
      calendar.createAllDayEvent(`${title} 마감 1주일전`, deadlineOneWeek);

      let deadlineFinal = new Date(deadlineOneWeek);
      deadlineFinal.setDate(deadlineOneWeek.getDate() + 8);
      calendar.createAllDayEvent(`${title} 마감 당일`, deadlineFinal);

      sheet.getRange(i + 1, 4).setValue("완료");
      count++;
    } catch (e) {
      console.log(`${i+1}행 등록 중 오류: ${e.message}`);
    }
  }
  ui.alert(`✅ 총 ${count}건의 취재 세트가 수정된 마감 알람 기준(+8일 / +16일)으로 등록되었습니다.`);
}

/**
 * 매일 새벽 자동 실행 트리거 생성
 */
function createJapanTripTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if(t.getHandlerFunction() === 'checkAndRegisterSchedules') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkAndRegisterSchedules').timeBased().everyDays(1).atHour(4).create();
  SpreadsheetApp.getUi().alert("매일 새벽 4시 자동 확인 설정 완료!");
}
