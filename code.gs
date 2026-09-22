/**
 * 하주초 연수 이수증 수합 시스템 - Backend
 * Model: gemini-2.5-flash
 */

const FOLDER_NAME = "연수 이수증";
const TEMPLATE_SHEET_NAME = "연수 수합(서식)";

function doGet(e) {
  try {
    // GET API 지원 (테스트 및 브라우저 직접 조회 호환)
    if (e && e.parameter && e.parameter.functionName) {
      const fn = e.parameter.functionName;
      const args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      const functionMap = {
        getTrainingList: getTrainingList,
        getApiKey: getApiKey,
        saveApiKey: saveApiKey,
        getParticipantList: getParticipantList,
        findParticipantsByName: findParticipantsByName,
        checkSubmissionsByIdentity: checkSubmissionsByIdentity
      };
      if (typeof functionMap[fn] === 'function') {
        const res = functionMap[fn].apply(null, args);
        return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: res }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    const template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_SHEET_NAME);
    if (!template) {
      return HtmlService.createHtmlOutput(`
        <div style="font-family: sans-serif; padding: 20px; text-align: center;">
          <h2 style="color: #ef4444;">설정 오류</h2>
          <p>[${TEMPLATE_SHEET_NAME}] 시트를 찾을 수 없습니다. 스프레드시트 설정을 확인해주세요.</p>
        </div>
      `).setTitle('오류 - 하주초 연수 이수증 수합');
    }

    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('하주초 연수 이수증 수합')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return HtmlService.createHtmlOutput("시스템 오류: " + e.toString());
  }
}

/**
 * 🚀 API Handler for Vercel/External Integration
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const functionName = params.functionName;
    const args = params.args || [];
    
    // 허용된 함수 매핑 (Apps Script V8 환경에서 this[functionName]이 undefined가 되는 현상 방지)
    const functionMap = {
      getTrainingList: getTrainingList,
      getApiKey: getApiKey,
      saveApiKey: saveApiKey,
      analyzeCertificate: analyzeCertificate,
      addTraining: addTraining,
      submitCertificate: submitCertificate,
      getParticipantList: getParticipantList,
      updateTraining: updateTraining,
      findParticipantsByName: findParticipantsByName,
      checkSubmissionsByIdentity: checkSubmissionsByIdentity
    };
    
    if (typeof functionMap[functionName] === 'function') {
      const result = functionMap[functionName].apply(null, args);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: result
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error(`Unauthorized or non-existent function: [${functionName}]`);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 관리자용: API 키 저장
 */
function saveApiKey(key) {
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  return "API 키가 저장되었습니다.";
}

function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || "";
}

/**
 * 등록된 연수 목록 가져오기 (최적화: CacheService + 일괄 셀 읽기 + 폴더URL 셀 캐시)
 */
function getTrainingList() {
  // ① 60초 캐시 확인 (반복 호출 방지)
  const cache = CacheService.getScriptCache();
  const cached = cache.get('trainingList');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const list = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name !== "메인" && !name.includes("설정") && !name.includes(TEMPLATE_SHEET_NAME)) {
      // ② 메타데이터 셀을 한 번에 일괄 읽기 (K1:P1 범위)
      const meta = sheet.getRange("A3:O3").getValues(); // 안내사항
      const row1 = sheet.getRange("K1:P1").getValues()[0]; // K:마감, L:-, M:담당자, N:-, O:대상, P:폴더URL캐시
      
      const deadline = row1[0]; // K1
      const manager = sheet.getRange("M1").getValue() || "미정";
      const target = sheet.getRange("O1").getValue() || "전 교직원";
      const notice = sheet.getRange("A3").getValue();

      let deadlineStr = "없음";
      try {
        if (deadline) {
          const d = new Date(deadline);
          if (!isNaN(d.getTime())) {
            deadlineStr = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
          }
        }
      } catch (e) {}

      // ③ 폴더 URL: P1 셀에 캐시된 값 우선 사용, 없으면 Drive 검색 후 저장
      let folderUrl = row1[5]; // P1
      if (!folderUrl) {
        folderUrl = getStorageFolder(name).getUrl();
        sheet.getRange("P1").setValue(folderUrl); // 다음 호출을 위해 저장
      }

      list.push({
        title: name,
        deadlineStr: deadlineStr,
        sheetUrl: `${ss.getUrl()}#gid=${sheet.getSheetId()}`,
        folderUrl: folderUrl,
        notice: notice,
        manager: manager,
        target: target
      });
    }
  });

  // ④ 결과를 60초간 캐싱
  try { cache.put('trainingList', JSON.stringify(list), 60); } catch(e) {}
  return list;
}

/**
 * 새 연수 추가: 템플릿 시트 복제 후 이름 변경
 */
function addTraining(title, deadline, notice, manager, target) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  
  if (ss.getSheetByName(title)) {
    throw new Error("이미 동일한 이름의 연수가 존재합니다.");
  }
  
  const newSheet = template.copyTo(ss).setName(title);
  // 마감 기한 저장 (J1: 라벨, K1: 데이터)
  newSheet.getRange("J1").setValue("마감 기한");
  newSheet.getRange("K1").setValue(deadline);
  newSheet.getRange("A1").setValue(title + " 연수 이수 결과");
  
  if (notice) {
    newSheet.getRange("A3").setValue(notice);
  }
  
  newSheet.getRange("L1").setValue("담당자");
  newSheet.getRange("M1").setValue(manager || "");
  newSheet.getRange("N1").setValue("대상");
  newSheet.getRange("O1").setValue(target || "전 교직원");

  // 폴더 URL을 P1에 미리 저장 (getTrainingList 속도 최적화)
  const folderUrl = getStorageFolder(title).getUrl();
  newSheet.getRange("P1").setValue(folderUrl);

  // 캐시 무효화
  CacheService.getScriptCache().remove('trainingList');
  
  return `[${title}] 연수가 등록되었습니다.`;
}

/**
 * 연수 정보 수정: 기존 시트의 메타데이터 및 이름을 업데이트
 */
function updateTraining(oldTitle, newTitle, deadline, notice, manager, target) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(oldTitle);
  
  if (!sheet) throw new Error(`[${oldTitle}] 연수 시트를 찾을 수 없습니다.`);
  
  // 제목이 바뀐 경우 시트 이름 변경 + 폴더 URL 재저장
  if (oldTitle !== newTitle) {
    if (ss.getSheetByName(newTitle)) throw new Error("이미 동일한 이름의 연수가 존재합니다.");
    sheet.setName(newTitle);
    // 폴더명이 바뀌었으므로 URL 재저장
    const folderUrl = getStorageFolder(newTitle).getUrl();
    sheet.getRange("P1").setValue(folderUrl);
  }
  
  // 메타데이터 업데이트
  sheet.getRange("K1").setValue(deadline);
  sheet.getRange("A1").setValue(newTitle + " 연수 이수 결과");
  sheet.getRange("A3").setValue(notice || "");
  sheet.getRange("M1").setValue(manager || "");
  sheet.getRange("O1").setValue(target || "전 교직원");

  // 캐시 무효화
  CacheService.getScriptCache().remove('trainingList');
  
  return `[${newTitle}] 연수 정보가 수정되었습니다.`;
}

/**
 * Gemini API를 활용한 이수증 분석 최적화
 */
function analyzeCertificate(base64Data, fileName) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error("GEMINI_API_KEY가 설정되지 않았습니다.");
  
  const mimeType = fileName.toLowerCase().endsWith(".pdf") ? "application/pdf" : "image/jpeg";
  const binaryData = base64Data.split(",")[1];
  
  // v1beta 모델 사용 (안정성 및 최신 기능 지원)
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;
  
  const prompt = `이 연수 이수증에서 다음 정보를 추출해서 JSON 형식으로 응답해줘.
1. name: 성명
2. courseName: "과정명" 항목의 전체 텍스트를 정확히 추출해줘. (주의: "연수 종류" 항목과 혼동하지 마. 과정명은 보통 매우 깁니다.)
3. certNo: 문서 상단이나 하단에 위치한 '제'와 '호'를 포함한 전체 이수번호 (예: 제 경북교연-2026-55579 호)
4. hours: "이수 시간" 항목의 텍스트를 "19시간 (1140분)"과 같이 시간과 분이 모두 포함된 형태로 추출해줘.
5. institution: 근무기관

마크다운 없이 순수 JSON만 반환: {"name":"","courseName":"","certNo":"","hours":"","institution":""}`;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: binaryData } }
      ]
    }],
    generationConfig: {
      maxOutputTokens: 300,
      temperature: 0.1,
      topP: 1
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  let retryCount = 0;
  let result = null;
  
  while (retryCount < 2) {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const content = response.getContentText();
    result = JSON.parse(content);
    
    // 할당량 초과(429) 재시도 logic
    if (response.getResponseCode() === 429 || (result.error && result.error.code === 429)) {
      Utilities.sleep(6000 + (retryCount * 2000)); // 점진적으로 대기 시간 증가
      retryCount++;
      continue;
    }
    break;
  }
  
  const defaultData = { "isCertificate": false, "name": "인식 실패", "courseName": "", "certNo": "", "hours": "", "institution": "" };

  if (result && result.candidates && result.candidates[0].content) {
    let text = result.candidates[0].content.parts[0].text;
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      try {
        const parsed = JSON.parse(jsonMatch[0]);
        // [하이브리드 핵심] 성명이나 과정명이 하나라도 추출되었다면 이수증(true)으로 자동 간주
        const hasAnyData = !!((parsed.name && parsed.name !== "인식 실패") || parsed.courseName || parsed.certNo);
        return { ...defaultData, ...parsed, "isCertificate": hasAnyData };
      } catch (e) {
        console.error("JSON 파싱 에러:", text);
      }
    }
  } else if (result.error) {
    return { ...defaultData, "isCertificate": false, "name": "API 오류: " + result.error.message };
  }
  
  return defaultData;
}

/**
 * 연수생 목록 가져오기 (학번/직위, 성명, 행번호)
 */
function getParticipantList(trainingTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(trainingTitle);
  if (!sheet) return [];

  // 실제 마지막 행까지만 읽기 (100행 고정 대신 동적 감지)
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  const values = sheet.getRange(4, 2, lastRow - 3, 2).getValues(); // B4:C{lastRow}
  const list = [];
  for (let i = 0; i < values.length; i++) {
    if (values[i][1]) {
      list.push({
        rowIndex: i + 4,
        classTitle: values[i][0],
        name: values[i][1]
      });
    }
  }
  return list;
}

/**
 * 최종 제출: 시트 기록 및 파일 저장
 */
function submitCertificate(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(data.trainingTitle);
  if (!sheet) throw new Error("해당 연수 시트를 찾을 수 없습니다.");
  
  let finalRowIndex = data.rowIndex;
  let finalName = data.name;
  let finalClassTitle = "";

  // rowIndex가 없는 경우에만 검색 수행
  if (!finalRowIndex) {
    const range = sheet.getRange("B4:C100").getValues();
    let matches = [];
    for (let i = 0; i < range.length; i++) {
      if (range[i][1] == data.name) {
        matches.push({
          rowIndex: i + 4,
          classTitle: range[i][0],
          name: range[i][1]
        });
      }
    }
    
    if (matches.length === 0) {
      throw new Error(`성명 [${data.name}]을(를) 직원 명단에서 찾을 수 없습니다.`);
    }
    
    if (matches.length === 1) {
      finalRowIndex = matches[0].rowIndex;
      finalClassTitle = matches[0].classTitle;
    } else {
      return { status: 'multiple', matches: matches };
    }
  }
  
  // 최종 기록
  // 과정명(D열/4번: D-F 병합) - 글씨크기 10, 왼쪽 정렬
  sheet.getRange(finalRowIndex, 4)
    .setValue(data.courseName || "")
    .setFontSize(10)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  
  // 이수번호(G열/7번: G-I 병합) - 글씨크기 10, 가운데 정렬
  sheet.getRange(finalRowIndex, 7)
    .setValue(data.certNo || "")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  saveFileToDrive(data.fileData, data.fileName, data.trainingTitle, data.name);
  
  return `${finalClassTitle ? '[' + finalClassTitle + '] ' : ''}${data.name} 님의 제출이 완료되었습니다.`;
}

function saveFileToDrive(base64Data, fileName, trainingTitle, userName) {
  const folder = getStorageFolder(trainingTitle);
  
  // 갱신 시 기존 파일 중복 저장을 방지하기 위해 이수증(userName) 파일이 이미 존재한다면 휴지통으로 이동
  try {
    const files = folder.getFiles();
    const searchPattern = `이수증(${userName})`;
    const toTrash = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().startsWith(searchPattern)) {
        toTrash.push(file);
      }
    }
    toTrash.forEach(file => {
      file.setTrashed(true);
    });
  } catch (fileErr) {
    console.error("기존 파일 삭제 중 오류 발생 (무시하고 진행): " + fileErr.toString());
  }

  const contentType = base64Data.split(",")[0].split(":")[1].split(";")[0];
  const binaryData = Utilities.base64Decode(base64Data.split(",")[1]);
  
  // 파일 확장자 추출
  const ext = fileName.includes('.') ? fileName.split('.').pop() : (contentType.includes('/') ? contentType.split('/')[1] : 'pdf');
  
  // 파일명 형식: 이수증(성명).확장자
  const newFileName = `이수증(${userName}).${ext}`;
  
  const blob = Utilities.newBlob(binaryData, contentType, newFileName);
  folder.createFile(blob);
}

function getStorageFolder(trainingTitle) {
  let rootFolder;
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) rootFolder = folders.next();
  else rootFolder = DriveApp.createFolder(FOLDER_NAME);
  
  const subFolders = rootFolder.getFoldersByName(trainingTitle);
  if (subFolders.hasNext()) return subFolders.next();
  else return rootFolder.createFolder(trainingTitle);
}

function getStorageFolderUrl() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next().getUrl();
  return "#";
}

/**
 * 성명으로 본인(들) 찾기 (동명이인 포함)
 */
function findParticipantsByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const participants = [];
  const seen = new Set();
  
  // 첫 번째 연수 시트(템플릿 등 제외)에서 명단을 가져옴
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetName !== "메인" && !sheetName.includes("설정") && !sheetName.includes(TEMPLATE_SHEET_NAME)) {
      const values = sheet.getRange("B4:C100").getValues();
      for (let i = 0; i < values.length; i++) {
        const classTitle = values[i][0];
        const participantName = values[i][1];
        if (participantName === name) {
          const key = classTitle + "|" + participantName;
          if (!seen.has(key)) {
            participants.push({ classTitle, name: participantName });
            seen.add(key);
          }
        }
      }
      // 명단은 모든 시트가 동일하다고 가정하므로 하나만 확인 후 중단
      break;
    }
  }
  return participants;
}

/**
 * 특정 사용자의 모든 연수 이수 현황 확인
 */
function checkSubmissionsByIdentity(name, classTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const submittedTrainings = [];
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName !== "메인" && !sheetName.includes("설정") && !sheetName.includes(TEMPLATE_SHEET_NAME)) {
      const values = sheet.getRange("B4:I100").getValues(); 
      for (let i = 0; i < values.length; i++) {
        const rowClass = values[i][0];
        const rowName = values[i][1];
        
        if (rowName === name && rowClass === classTitle) {
          const hasCourse = values[i][2] && values[i][2].toString().trim() !== "";
          const hasCertNo = values[i][5] && values[i][5].toString().trim() !== "";
          
          if (hasCourse || hasCertNo) {
            submittedTrainings.push(sheetName);
          }
          break;
        }
      }
    }
  });
  
  return submittedTrainings;
}
