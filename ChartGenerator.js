/**
 * ===============================================================================
 * [차트 자동화] 2026_보고서ver2 컬럼 매핑 완벽 수정 버전
 * ===============================================================================
 */

function generateBatchChartsFromSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const activeSheet = ss.getActiveSheet();

  // 1. 시트 확인
  if (activeSheet.getName() !== CONFIG.SHEETS.CHART_SETTING) {
    ui.alert(`⚠️ [${CONFIG.SHEETS.CHART_SETTING}] 시트에서 영역을 선택한 후 실행해 주세요.`);
    return;
  }

  // 2. 선택된 범위 확인
  const activeRange = activeSheet.getActiveRange();
  if (!activeRange) {
    ui.alert("⚠️ 작업할 행이나 셀을 먼저 선택해 주세요.");
    return;
  }

  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();

  if (startRow < 2) {
    ui.alert("⚠️ 헤더(1행)를 제외한 2행 이후의 데이터 영역을 선택해 주세요.");
    return;
  }

  // 3. 차트 백엔드 데이터 시트 준비
  let dataSheet = ss.getSheetByName(CONFIG.SHEETS.CHART_DATA);
  if (!dataSheet) {
    dataSheet = ss.insertSheet(CONFIG.SHEETS.CHART_DATA);
  }
  
  if (dataSheet.getLastRow() === 0) {
    const headers = ["취재일", "클라이언트", "블로거명", "구분"];
    for (let d = 1; d <= 14; d++) headers.push(`${d}日`);
    headers.push("실제 평균");
    dataSheet.appendRow(headers);
    dataSheet.getRange(1, 1, 1, headers.length).setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
  }

  // 4. 최상위 루트 폴더 [14일차트] 확인 및 생성
  const rootFolder = getOrCreateFolder("14일차트");

  let processedCount = 0;

  // 5. 선택된 행 순회
  for (let i = 0; i < numRows; i++) {
    const currentRowIndex = startRow + i;
    // 2026_보고서ver2의 최소 26열까지 넉넉하게 읽음
    const rowValues = activeSheet.getRange(currentRowIndex, 1, 1, 26).getValues()[0];

    // D열(인덱스 3): 취재일, E열(인덱스 4): 클라이언트명
    const rawDate = rowValues[3];
    const tripDateFormatted = formatToYYMMDD(rawDate);
    const clientName = rowValues[4] ? rowValues[4].toString().trim() : "기타클라이언트";

    // --- [타겟 1: 블로거 1 처리 (F:이름, J:PV, K:진행상태, L:구글드라이브링크)] ---
    const name1 = rowValues[5] ? rowValues[5].toString().trim() : "";   // F열 (인덱스 5)
    const pv1 = Number(rowValues[9]);                                   // J열: PV1 (인덱스 9)
    const status1 = rowValues[10] ? rowValues[10].toString().trim() : ""; // K열: 진행상태1 (인덱스 10)

    if (name1 && !isNaN(pv1) && pv1 > 0 && status1 !== "✅ 완료") {
      const clientFolder = getOrCreateSubFolder(rootFolder, clientName);
      const filePrefix = tripDateFormatted ? `${tripDateFormatted}_${name1}` : name1;
      
      // 세로 버퍼는 22열(V), 23열(W) 사용
      processSingleBloggerChart(dataSheet, tripDateFormatted, clientName, name1, "블로거1", pv1, clientFolder, filePrefix, 22);
      
      activeSheet.getRange(currentRowIndex, 11).setValue("✅ 완료");          // K열(11): 진행상태1
      activeSheet.getRange(currentRowIndex, 12).setValue(clientFolder.getUrl()); // L열(12): 구글드라이브링크1
      processedCount++;
    }

    // --- [타겟 2: 블로거 2 처리 (Q:이름, U:PV, V:진행상태, W:구글드라이브링크)] ---
    const name2 = rowValues[16] ? rowValues[16].toString().trim() : ""; // Q열: 이름2 (인덱스 16)
    const pv2 = Number(rowValues[20]);                                   // U열: PV2 (인덱스 20)
    const status2 = rowValues[21] ? rowValues[21].toString().trim() : ""; // V열: 진행상태2 (인덱스 21)

    if (name2 && !isNaN(pv2) && pv2 > 0 && status2 !== "✅ 완료") {
      const clientFolder = getOrCreateSubFolder(rootFolder, clientName);
      const filePrefix = tripDateFormatted ? `${tripDateFormatted}_${name2}` : name2;
      
      // 세로 버퍼는 24열(X), 25열(Y) 사용
      processSingleBloggerChart(dataSheet, tripDateFormatted, clientName, name2, "블로거2", pv2, clientFolder, filePrefix, 24);
      
      activeSheet.getRange(currentRowIndex, 22).setValue("✅ 완료");          // V열(22): 진행상태2
      activeSheet.getRange(currentRowIndex, 23).setValue(clientFolder.getUrl()); // W열(23): 구글드라이브링크2
      processedCount++;
    }
  }

  SpreadsheetApp.flush();

  if (processedCount > 0) {
    ui.alert(`🎉 작업 완료!\n총 ${processedCount}명의 차트가 T열에 생성되고 구글 드라이브에 저장되었습니다.`);
  } else {
    ui.alert("💡 처리할 대상이 없거나, 이미 모든 대상이 '✅ 완료' 상태입니다.");
  }
}

/**
 * [단일 블로거 처리] T열(20열) 차트 삽입 및 일본어(日付, 1日~14日) 적용
 */
function processSingleBloggerChart(dataSheet, tripDate, clientName, bloggerName, targetType, targetAvg, targetFolder, filePrefix, targetBufferCol) {
  const totalDays = 14;
  const targetTotal = targetAvg * totalDays;

  // 1. 평균값 기준 ±10% 범위 난수 생성
  const fluctuationRatio = 0.10;
  let rawValues = [];
  for (let d = 0; d < totalDays; d++) {
    const randomPercent = (Math.random() * (fluctuationRatio * 2)) - fluctuationRatio;
    const val = Math.max(1, targetAvg * (1 + randomPercent));
    rawValues.push(val);
  }

  // 2. 정규화 (14일 평균 일치)
  const rawSum = rawValues.reduce((acc, cur) => acc + cur, 0);
  let normalizedValues = rawValues.map(v => Math.round((v / rawSum) * targetTotal));
  
  const currentSum = normalizedValues.reduce((acc, cur) => acc + cur, 0);
  normalizedValues[totalDays - 1] += (targetTotal - currentSum);

  // 3. A열 기준 순차 행 계산 및 가로 데이터 기록 (A~S열)
  const aColValues = dataSheet.getRange("A:A").getValues();
  let nextRow = 1;
  while (nextRow <= aColValues.length && aColValues[nextRow - 1][0] !== "") {
    nextRow++;
  }

  const rowData = [tripDate, clientName, bloggerName, targetType, ...normalizedValues];
  dataSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
  dataSheet.getRange(nextRow, 19).setFormula(`=AVERAGE(E${nextRow}:R${nextRow})`); // S열 수식
  dataSheet.getRange(nextRow, 5, 1, 15).setNumberFormat("#,##0");

  // 4. 전용 세로 데이터 버퍼 채우기 (일자 ➔ 日付, 1일 ➔ 1日)
  const chartDataArray = [["日付", "閲覧数"]];
  for (let d = 0; d < totalDays; d++) {
    chartDataArray.push([`${d + 1}日`, normalizedValues[d]]);
  }
  
  const bufferRange = dataSheet.getRange(1, targetBufferCol, 15, 2);
  bufferRange.setValues(chartDataArray);
  dataSheet.getRange(2, targetBufferCol + 1, 14, 1).setNumberFormat("#,##0");
  SpreadsheetApp.flush();

  // 5. Y축 안전 범위 계산
  const minVal = Math.min(...normalizedValues);
  const maxVal = Math.max(...normalizedValues);
  const rangeSpan = (maxVal - minVal) || (targetAvg * 0.1);
  const yAxisMin = Math.max(0, Math.floor((minVal - rangeSpan * 1.5) / 100) * 100);
  const yAxisMax = Math.ceil((maxVal + rangeSpan * 1.5) / 100) * 100;

  // 6. 차트 생성 및 T열(20번째 열) 삽입
  const sheetChart = dataSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(bufferRange)
    .setPosition(nextRow, 20, 0, 0) // T열 배치
    .setOption('legend', { position: 'none' })
    .setOption('curveType', 'function')
    .setOption('colors', ['#2563eb'])
    .setOption('pointSize', 6)
    .setOption('series', {
      0: {
        dataLabel: 'value',
        dataLabelPlacement: 'top',
        annotations: {
          textStyle: { fontSize: 10, color: '#0f172a', bold: true },
          stem: { length: 5, color: 'transparent' },
          alwaysOutside: true
        }
      }
    })
    .setOption('chartArea', { left: '8%', top: '15%', width: '88%', height: '70%' })
    .setOption('hAxis', {
      slantedText: false,
      textStyle: { fontSize: 11, color: '#334155' }
    })
    .setOption('vAxis', {
      viewWindow: { min: yAxisMin, max: yAxisMax },
      format: '#,###',
      textStyle: { fontSize: 11, color: '#334155' },
      gridlines: { color: '#f1f5f9' }
    })
    .setOption('width', 850)
    .setOption('height', 380)
    .build();

  dataSheet.insertChart(sheetChart);

  // 7. PNG 이미지 추출 및 드라이브 저장
  const fileName = `${filePrefix}_14일차트.png`;
  const chartBlob = sheetChart.getAs('image/png').setName(fileName);

  const existingFiles = targetFolder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  targetFolder.createFile(chartBlob);
}

/**
 * [날짜 포맷 헬퍼] 다양한 형태의 날짜를 6자리 'YYMMDD'로 변환
 */
function formatToYYMMDD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyMMdd");
  }
  
  let str = val.toString().trim();
  let parts = str.match(/(\d+)/g);
  if (!parts) return str;

  if (parts.length === 3) {
    let y = parts[0];
    let m = parts[1].padStart(2, '0');
    let d = parts[2].padStart(2, '0');
    if (y.length === 4) y = y.substring(2);
    return `${y}${m}${d}`;
  } else if (parts.length === 2) {
    let currentYear = new Date().getFullYear().toString().substring(2);
    let m = parts[0].padStart(2, '0');
    let d = parts[1].padStart(2, '0');
    return `${currentYear}${m}${d}`;
  }
  return str.replace(/[^0-9]/g, '');
}

/**
 * [폴더 유틸] 최상위 폴더 가져오기 또는 생성
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * [폴더 유틸] 서브 폴더 가져오기 또는 생성
 */
function getOrCreateSubFolder(parentFolder, subFolderName) {
  const folders = parentFolder.getFoldersByName(subFolderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(subFolderName);
}
