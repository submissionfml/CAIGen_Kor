/**************************************************
 * CAIGen_Kor.gs
 **************************************************/

/********************
 아래에 json_files, gs_files 폴더가 있는 이 프로그램 루트 폴더 ID를 입력해야 합니다.
 구글드라이브 루트 폴더 인터넷 주소가 "https://drive.google.com/drive/u/0/folders/1l4ANmPeksTLNW4UMdnvleDzXK19FrfYP"면 이 주소의 가장 마지막 부부인 "1l4ANmPeksTLNW4UMdnvleDzXK19FrfYP"가 루트 폴더 ID입니다. 
 ********************/

var projectFolderId = "여기에 루트 폴더 ID를 넣어 주세요.";

/********************
 * Default Public Settings
 *
 * The values below are preconfigured for the public release.
 * Most users do not need to change them.
 ********************/

// The preprocess_to_caigen tool generates upload_data.json by default.
var sourceFiles = ["upload_data.json"];

// Spreadsheet filename to create inside the gs_files folder.
var outputSpreadsheetName = "annotation_workbook";

// Default folder names inside the project folder.
var dataFolderName = "json_files";
var outFolderName = "gs_files";

// Set true only for short debugging runs.
var debug = false;

/********************
 * Public Functions: Setup
 ********************/
function createAndWriteSheets() {
  createSheets();
  writeSheets();
}

function getProjectFolderOrThrow() {
  try {
    return DriveApp.getFolderById(projectFolderId);
  } catch (e) {
    throw new Error("구글드라이브 루트 폴더 id를 다시 확인해서 바르게 입력해 주세요.");
  }
}

function createSheets() {
  var projectFolder = getProjectFolderOrThrow();
  var outFolder = findChildFolderByName(projectFolder, outFolderName);
  cleanFolder(outFolder);

  // Set up a spreadsheet and sheets
  var ss = createSpreadsheet(outFolder);
  ss.rename(outputSpreadsheetName);
  Logger.log(`Created spreadsheet:`);
  Logger.log(ss.getUrl());

  var facesheet = ss.getSheetByName("시트1");
  facesheet.setName("facesheet");
  facesheet.setColumnWidth(1, 440);
  // Optional notes for annotators can be added here.
  // facesheet.getRange(1, 1).setValue('Field 1');
  // facesheet.getRange(2, 1).setValue('Field 2');

  sourceFiles.forEach((sourceFile) => {
    var sheetName = sourceFile.split(".")[0];
    var sheet = ss.insertSheet();
    sheet.setName(sheetName);
    sheet.setColumnWidth(1, 140);
    // Increase rows to avoid bugs
    sheet.insertRowsAfter(1000, 99000);
  });

  // CharOffset 시트 추가
  var charOffsetSheet = ss.insertSheet();
  charOffsetSheet.setName("CharOffset");
  charOffsetSheet.setColumnWidth(1, 400); // A열 너비 조정
  charOffsetSheet.insertRowsAfter(1000, 99000); // 약 100,000행 확보
}

/********************
 * Public Functions: Writing
 ********************/

/**
 * Writing notes
 * - For large jobs, let the spreadsheet finish filling before starting annotation.
 * - If execution stops midway, run writeSheetsResume().
 * - If you see "[SKIP] Another processNextItem execution is already running.", wait a moment and try again.
 * - Very long text may exceed a Google Sheets cell limit; such rows are marked as skipped.
 */

// These functions are designed to restart safely when Apps Script stops because of time limits.

function writeSheets() {
  // 수동 실행(writeSheets 클릭) 시에는 항상 처음부터 시작
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("nextSentIndex", 0);
  clearAllWriteRowPointers();
  deleteAllTriggers();

  // Filling many cells takes time while the script has a time limit of 6 minutes, so measures against TimeOutError are necessary
  // In this recursive function, we:
  // - build a list of sentences to be written, i.e., sentList
  // - write sentences one by one, updating sentIndex
  // - if the time limit is reached, save the state and set a trigger for the next iteration
  processNextItem();
}

function writeSheetsResume() {
  // 저장된 nextSentIndex부터 이어서 실행하고 싶을 때 수동 호출
  deleteAllTriggers();
  processNextItem();
}

function restartWriteSheets() {
  // 기존 이름 호환용: 강제로 처음부터 다시 시작
  writeSheets();
}

/********************
 * Main Processing (Internal)
 ********************/

function processNextItem() {
  // 동시 실행 방지: 같은 시점에 여러 실행이 같은 시트에 쓰지 못하도록 잠금
  var scriptLock = LockService.getScriptLock();
  if (!scriptLock.tryLock(1000)) {
    Logger.log("[SKIP] Another processNextItem execution is already running.");
    return;
  }

  try {
    // The definitions of `projectFolderId`, `dataFolderName`, and `outFolderName` are in CreateFiles.gs
    var projectFolder = getProjectFolderOrThrow();
    var dataFolder = findChildFolderByName(projectFolder, dataFolderName);
    var outFolder = findChildFolderByName(projectFolder, outFolderName);
    var ssFile = findFileByName(outFolder, outputSpreadsheetName);
    var ss = SpreadsheetApp.openById(ssFile.getId());

    // Each sentence in sentList is represented as an object with the keys: dataName, sentData
    // dataName is the name of the data file (without the extension), e.g., 'sheet_1'
    // sentData has the keys: id, target_sentence, context_left, context_right, tokens
    var sentList = buildSentenceList(sourceFiles, dataFolder);
    if (debug) {
      sentList = sentList.slice(0, 2);
    }

    var scriptProperties = PropertiesService.getScriptProperties();
    var startIndex = parseInt(scriptProperties.getProperty("nextSentIndex"), 10);
    if (isNaN(startIndex) || startIndex < 0) {
      startIndex = 0;
    }
    // By setting maxTime at 4.75 min, we will have 1.25 minute for setting a trigger
    // This sort of margin is important to avoid TimeOutError
    var maxTime = 4.75 * 60 * 1000;
    var startTime = new Date().getTime();

    // CharOffset 시트 가져오기 또는 생성
    var charOffsetSheet = ss.getSheetByName("CharOffset");
    if (!charOffsetSheet) {
      charOffsetSheet = ss.insertSheet();
      charOffsetSheet.setName("CharOffset");
      charOffsetSheet.setColumnWidth(1, 400); // A열 너비 조정
      charOffsetSheet.insertRowsAfter(1000, 99000); // 약 100,000행 확보
    }
    
    // 전역 문장 번호 추적 (모든 시트에 걸쳐 1, 2, 3... 순서로 증가)
    // CharOffset 시트의 A열에서 기존 문장 수 확인
    // 주의: A열 셀은 수식이 들어가지만, 체크 전에는 결과값이 "" 이라 getValues()로는 비어있는 것처럼 보일 수 있음.
    // 따라서 값(getValues)뿐 아니라 수식(getFormulas) 존재 여부도 함께 확인해야 타임아웃 재실행 시 덮어쓰기를 막을 수 있다.
    var globalSentenceCount = 0;
    var lastRow = charOffsetSheet.getLastRow();
    if (lastRow > 0) {
      var aColValues = charOffsetSheet.getRange(1, 1, lastRow, 1).getValues();
      var aColFormulas = charOffsetSheet.getRange(1, 1, lastRow, 1).getFormulas();
      for (var r = 0; r < lastRow; r++) {
        var hasValue = aColValues[r][0] !== "" && aColValues[r][0] !== null;
        var hasFormula = aColFormulas[r][0] !== "" && aColFormulas[r][0] !== null;
        if (hasValue || hasFormula) {
          globalSentenceCount = r + 1; // 1-based 행 번호
        }
      }
    }
    
    // 실행 시작 전에 남은 문장 수 기준으로 시트 행 수를 한 번에 확장
    // (행 부족으로 중간에 1번 문장으로 돌아가거나 쓰기 실패하는 문제 방지)
    preallocateRowsForRun(ss, sentList, startIndex, charOffsetSheet);

    // 각 시트별로 문장 번호 추적 (로컬 용도)
    var sheetSentenceCount = {};
    
    // Iteratively write sentences till the time limit is reached
    for (var i = startIndex; i < sentList.length; i++) {
      var dataName = sentList[i].dataName;
      
      // 각 시트별 문장 번호 계산 (1부터 시작, 로컬 용도)
      if (!sheetSentenceCount.hasOwnProperty(dataName)) {
        sheetSentenceCount[dataName] = 0;
      }
      sheetSentenceCount[dataName]++;
      var sentenceNumber = sheetSentenceCount[dataName];
      
      // 전역 문장 번호 증가 (CharOffset 시트의 A열에 저장할 행 번호)
      globalSentenceCount++;
      
      // 시트가 이미 존재하면 append = true, 없으면 false
      var append = ss.getSheetByName(dataName) !== null;
      try {
        writeSheet(ss, dataName, sentList[i].sentData, append, sentenceNumber, globalSentenceCount, charOffsetSheet);
      } catch (e) {
        recordSentenceWriteError(ss, dataName, sentList[i].sentData, e);
        Logger.log(`[에러] 문장 쓰기 실패 (계속 진행): ${sentList[i].sentData && sentList[i].sentData.id ? sentList[i].sentData.id : '(unknown id)'} / ${e}`);
      }

      if (new Date().getTime() - startTime > maxTime) {
        // Save the state and setup a trigger for the next execution
        scriptProperties.setProperty("nextSentIndex", i + 1);
        ScriptApp.newTrigger("processNextItem")
          .timeBased()
          .after(1 * 1000)
          .create();
        return;
      }
    }

    // Report the completion
    var facesheet = ss.getSheetByName("facesheet");
    if (facesheet) {
      var row = Math.max(facesheet.getLastRow() + 1, 101);
      facesheet
        .getRange(row, 1)
        .setValue(`Completed writing ${sentList.length} sentences`);
    }
  } finally {
    try {
      scriptLock.releaseLock();
    } catch (e) {
      // ignore release errors
    }
  }
}

/**
 * JSON/JSONL 파일 파싱 함수
 * @param {string} content - 파일 내용
 * @param {string} fileName - 파일 이름 (확장자 판별용)
 * @returns {Array} 파싱된 객체 배열
 */
function parseJsonOrJsonl(content, fileName) {
  var extension = fileName.split(".").pop().toLowerCase();
  
  if (extension === "jsonl") {
    // JSONL: 한 줄씩 JSON.parse
    var lines = content.split("\n");
    var data = [];
    lines.forEach(function(line) {
      line = line.trim();
      if (line.length > 0) {
        data.push(JSON.parse(line));
      }
    });
    return data;
  } else {
    // JSON: 기존 방식 (배열로 가정)
    return JSON.parse(content);
  }
}

function buildSentenceList(inFiles, dataFolder) {
  var sentences = [];
  inFiles.forEach((inFile) => {
    var dataFile = findFileByName(dataFolder, inFile);
    var dataName = dataFile.getName().split(".")[0];
    var content = dataFile.getBlob().getDataAsString();
    // JSON/JSONL 파서 사용
    var data = parseJsonOrJsonl(content, dataFile.getName());
    data.forEach((obj) => {
      sentences.push({ dataName: dataName, sentData: obj });
    });
  });

  return sentences;
}

function ensureSheetHasRows(sheet, targetRows) {
  var currentMaxRows = sheet.getMaxRows();
  if (targetRows <= currentMaxRows) return;

  var rowsToAdd = targetRows - currentMaxRows;
  sheet.insertRowsAfter(currentMaxRows, rowsToAdd);
}

function estimateRowsNeededForSentence(sentData) {
  // writeSheet()와 같은 레이아웃 상수 사용
  var wordPerLine = 18;
  var maxMwePerSent = 9;
  var bandWidth = maxMwePerSent + 5; // 14

  var tokenCount = 0;
  if (sentData && sentData.tokens && sentData.tokens.length) {
    tokenCount = sentData.tokens.length;
  }

  // 토큰이 없어도 최소 1 band 레이아웃은 확보
  var numBands = Math.max(1, Math.ceil(tokenCount / wordPerLine));

  // 다음 문장 시작 행 증가량 기준(문장 블록 + 문장 간 공백 1행 포함)
  // 1-band일 때 약 30행, band 하나 늘 때마다 14행 증가
  return 16 + bandWidth * numBands;
}

function preallocateRowsForRun(ss, sentList, startIndex, charOffsetSheet) {
  var remainingRowsBySheet = {};

  for (var i = startIndex; i < sentList.length; i++) {
    var item = sentList[i];
    if (!remainingRowsBySheet[item.dataName]) {
      remainingRowsBySheet[item.dataName] = 0;
    }
    remainingRowsBySheet[item.dataName] += estimateRowsNeededForSentence(item.sentData);
  }

  // 데이터 시트별로 현재 마지막 사용 행 + 남은 예상 행 수만큼 한 번에 확장
  for (var dataName in remainingRowsBySheet) {
    var sheet = ss.getSheetByName(dataName);
    if (!sheet) continue;

    var currentLastRow = sheet.getLastRow();
    var targetRows = currentLastRow + remainingRowsBySheet[dataName] + 100; // 여유 버퍼
    ensureSheetHasRows(sheet, targetRows);
  }

  // CharOffset 시트는 문장당 1행 사용 (+여유 버퍼)
  if (charOffsetSheet) {
    var targetCharOffsetRows = sentList.length + 100;
    ensureSheetHasRows(charOffsetSheet, targetCharOffsetRows);
  }
}


function getWriteRowPointerKey(ssId, dataName) {
  return "nextWriteRow:" + ssId + ":" + dataName;
}

function getNextWriteRowPointer(ss, dataName, sheet) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var key = getWriteRowPointerKey(ss.getId(), dataName);
  var raw = scriptProperties.getProperty(key);
  var parsed = parseInt(raw, 10);
  if (!isNaN(parsed) && parsed > 0) {
    return parsed;
  }

  var fallback = 1;
  if (sheet) {
    var lastRow = sheet.getLastRow();
    fallback = lastRow > 0 ? lastRow + 2 : 1;
  }
  scriptProperties.setProperty(key, String(fallback));
  return fallback;
}

function setNextWriteRowPointer(ss, dataName, nextRow) {
  var parsed = parseInt(nextRow, 10);
  if (isNaN(parsed) || parsed < 1) return;
  var scriptProperties = PropertiesService.getScriptProperties();
  var key = getWriteRowPointerKey(ss.getId(), dataName);
  scriptProperties.setProperty(key, String(parsed));
}

function clearAllWriteRowPointers() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var all = scriptProperties.getProperties();
  for (var key in all) {
    if (key.indexOf("nextWriteRow:") === 0) {
      scriptProperties.deleteProperty(key);
    }
  }
}


function recordSentenceWriteError(ss, dataName, sentData, err) {
  var errMsg = err && err.message ? String(err.message) : String(err);
  var sentId = sentData && sentData.id ? sentData.id : "(unknown id)";
  var targetSheet = ss.getSheetByName(dataName);

  // facesheet 로그
  var facesheet = ss.getSheetByName("facesheet");
  if (facesheet) {
    var logRow = Math.max(facesheet.getLastRow() + 1, 101);
    facesheet.getRange(logRow, 1).setValue(`[ERROR] ${sentId} on ${dataName}: ${truncateForCell(errMsg, 450)}`);
  }

  // 별도 ErrorLog 시트에 상세 기록 (문장 처리 계속)
  var errorSheet = ss.getSheetByName("ErrorLog");
  if (!errorSheet) {
    errorSheet = ss.insertSheet();
    errorSheet.setName("ErrorLog");
    errorSheet.getRange(1, 1, 1, 5).setValues([["timestamp", "sheet", "sentence_id", "error", "note"]]);
    errorSheet.setFrozenRows(1);
  }
  var erow = errorSheet.getLastRow() + 1;
  errorSheet.getRange(erow, 1, 1, 5).setValues([[new Date(), dataName, sentId, truncateForCell(errMsg, 49000), isCellLimitError(errMsg) ? "Skipped sentence due to Google Sheets 50,000-char cell limit" : "Skipped sentence due to write error"]]);

  // 대상 시트에도 간단한 표시 남기기 (길이 안전한 텍스트만 사용)
  if (targetSheet) {
    var row = getNextWriteRowPointer(ss, dataName, targetSheet);
    ensureSheetHasRows(targetSheet, row + 2);
    targetSheet.getRange(row, 1).setValue(`[ERROR] ${sentId}`);
    targetSheet.getRange(row, 2).setValue(isCellLimitError(errMsg) ? "Skipped: exceeds Google Sheets 50,000 chars per cell" : "Skipped: write error");
    targetSheet.getRange(row + 1, 1).setValue("Sentence");
    targetSheet.getRange(row + 1, 2).setValue(truncateForCell((sentData && (sentData.target_sentence || sentData.text)) || "", 1000));
    targetSheet.getRange(row, 1, 2, 2).setBackground("#fff3f3");
    setNextWriteRowPointer(ss, dataName, row + estimateRowsNeededForSentence(sentData));
  }
}

function truncateForCell(text, maxLen) {
  var s = text == null ? "" : String(text);
  if (s.length <= maxLen) return s;
  return s.slice(0, Math.max(0, maxLen - 3)) + "...";
}

function isCellLimitError(message) {
  var s = String(message || "");
  return s.indexOf("50000") !== -1 || s.indexOf("최대 문자수") !== -1;
}

function detectOversizedCellInput(sentData) {
  var MAX_CELL_CHARS = 50000;
  if (!sentData) return null;

  function checkField(value, fieldName) {
    if (value === null || value === undefined) return null;
    var s = String(value);
    if (s.length > MAX_CELL_CHARS) {
      return { type: "field", field: fieldName, length: s.length };
    }
    return null;
  }

  var issue = null;
  issue = checkField(sentData.target_sentence, "target_sentence");
  if (issue) return issue;
  issue = checkField(sentData.text, "text");
  if (issue) return issue;
  issue = checkField(sentData.context_left, "context_left");
  if (issue) return issue;
  issue = checkField(sentData.context_right, "context_right");
  if (issue) return issue;

  if (sentData.tokens && sentData.tokens.length) {
    for (var i = 0; i < sentData.tokens.length; i++) {
      var tok = sentData.tokens[i] || {};
      issue = checkField(tok.surface, "tokens[" + i + "].surface");
      if (issue) {
        issue.type = "token";
        issue.tokenIndex = i;
        return issue;
      }
    }
  }

  return null;
}

function safePreviewText(value, maxLen) {
  var s = value == null ? "" : String(value);
  if (s.length <= maxLen) return s;
  return s.slice(0, Math.max(0, maxLen - 24)) + " ...(truncated " + s.length + ")";
}

function writeSkippedSentencePreviewBlock(sheet, row_i, sentData, issue, wordPerLine) {
  var maxMwePerSent = 9;
  var previewTokenLimit = 100;
  var sentId = sentData && sentData.id ? sentData.id : "(unknown id)";
  var tokens = (sentData && sentData.tokens) ? sentData.tokens : [];
  var previewCount = Math.min(previewTokenLimit, tokens.length);
  var previewBands = Math.max(1, Math.ceil(previewCount / wordPerLine));

  // Header / status block
  sheet.getRange(row_i, 1).setValue(sentId);
  sheet.getRange(row_i, 1, 1, wordPerLine + 1).setBackground("#ffeaea");
  sheet.getRange(row_i + 1, 1).setValue("Status");
  sheet.getRange(row_i + 1, 2).setValue("SKIPPED (cell > 50,000 chars)");
  sheet.getRange(row_i + 2, 1).setValue("Reason");
  sheet.getRange(row_i + 2, 2).setValue("Exceeded Google Sheets 50,000-char cell limit at " + issue.field + " (len=" + issue.length + ")");
  sheet.getRange(row_i + 3, 1).setValue("Preview");
  sheet.getRange(row_i + 3, 2).setValue("First " + previewCount + " tokens only");

  // Preview token grid (index row + token row repeated)
  var tokenStartRow = row_i + 4;
  for (var bandIdx = 0; bandIdx < previewBands; bandIdx++) {
    var bandRow = tokenStartRow + bandIdx * 2;
    var start = bandIdx * wordPerLine;
    var end = Math.min(start + wordPerLine, previewCount);
    var indices = [];
    var surfaces = [];
    for (var j = start; j < end; j++) {
      indices.push(j + 1);
      var tok = tokens[j] || {};
      surfaces.push(safePreviewText(tok.surface, 1000));
    }
    sheet.getRange(bandRow, 1).setValue(bandIdx === 0 ? "Tok#" : "");
    sheet.getRange(bandRow + 1, 1).setValue(bandIdx === 0 ? "Tok" : "");
    sheet.getRange(bandRow, 2, 1, indices.length).setValues([indices]);
    sheet.getRange(bandRow + 1, 2, 1, surfaces.length).setValues([surfaces]);
  }

  var afterPreviewRow = tokenStartRow + previewBands * 2;
  var targetText = (sentData && (sentData.target_sentence || sentData.text)) || "";
  sheet.getRange(afterPreviewRow, 1).setValue("Sentence Preview");
  sheet.getRange(afterPreviewRow, 2).setValue(safePreviewText(targetText, 1000));
  if (sentData && sentData.context_left) {
    sheet.getRange(afterPreviewRow + 1, 1).setValue("Context (Left) Preview");
    sheet.getRange(afterPreviewRow + 1, 2).setValue(safePreviewText(sentData.context_left, 1000));
  }
  if (sentData && sentData.context_right) {
    var crr = afterPreviewRow + (sentData.context_left ? 2 : 1);
    sheet.getRange(crr, 1).setValue("Context (Right) Preview");
    sheet.getRange(crr, 2).setValue(safePreviewText(sentData.context_right, 1000));
  }

  // Mild formatting for skipped block
  var blockEndRow = afterPreviewRow + (sentData && sentData.context_left && sentData.context_right ? 2 : (sentData && (sentData.context_left || sentData.context_right) ? 1 : 0));
  sheet.getRange(row_i, 1, blockEndRow - row_i + 1, 2).setBorder(true, true, true, true, false, false);
}

function writeSheet(ss, dataName, sentData, append = false, localSentenceNumber = null, globalSentenceNumber = null, charOffsetSheet = null) {
  // When append is true, open an existing sheet and start writing rows at the bottom. Otherwise, create a new sheet
  // localSentenceNumber: 각 시트 내에서의 문장 순서 (1부터 시작)
  // globalSentenceNumber: 모든 시트에 걸쳐 전역 문장 번호 (1부터 시작), CharOffset 시트의 A열에 출력할 행 번호로 사용
  // charOffsetSheet: CharOffset 시트 객체

  var colWordStart = 2;
  var wordPerLine = 18; // 형태소 15-20개/줄 (18로 설정)
  var maxMwePerSent = 9;
  // bandWidth: index(1) + word(1) + checkbox(9) + start(1) + end(1) + index_hidden(1) = 14
  var bandWidth = maxMwePerSent + 5; // 14
  // 헤더 섹션: ID(1) + 헤더(1) + MWE행들(9) + Sentence(1) + context_left(1) + context_right(1) = 14
  var headerSectionHeight = maxMwePerSent + 5; // 14

  if (!append) {
    var sheet = ss.insertSheet();
    sheet.setName(dataName);
    sheet.setColumnWidth(1, 140);
    // Insert rows, which is necessary to avoid bugs
    sheet.insertRowsAfter(1000, 99000);
    var row_i = 1;
  } else {
    var sheet = ss.getSheetByName(dataName);
    var row_i = getNextWriteRowPointer(ss, dataName, sheet);
  }

  // 문장 시작 행은 getLastRow()가 아니라 저장된 포인터를 기준으로 사용한다.
  var rowStart = row_i;

  // 입력 텍스트 중 Google Sheets 셀 한도(50,000자)를 넘는 값이 있으면,
  // 해당 문장을 산출물 시트에 SKIPPED 블록으로 기록하고 다음 문장으로 진행한다.
  var oversizedInputIssue = detectOversizedCellInput(sentData);
  if (oversizedInputIssue) {
    writeSkippedSentencePreviewBlock(sheet, row_i, sentData, oversizedInputIssue, wordPerLine);

    if (globalSentenceNumber !== null && globalSentenceNumber > 0 && charOffsetSheet) {
      try {
        charOffsetSheet.getRange(globalSentenceNumber, 1).clearDataValidations();
      } catch (e) {}
      try {
        charOffsetSheet.getRange(globalSentenceNumber, 1).setValue('[SKIPPED] ' + (sentData && sentData.id ? sentData.id : '(unknown id)') + ' / oversized cell');
      } catch (e) {}
    }

    setNextWriteRowPointer(ss, dataName, rowStart + estimateRowsNeededForSentence(sentData));
    return;
  }

  // ==========
  // 헤더 섹션 작성

  // Log
  var facesheet = ss.getSheetByName("facesheet");
  if (facesheet) {
    var row = Math.max(facesheet.getLastRow() + 1, 101);
    facesheet
      .getRange(row, 1)
      .setValue(`Writing ${sentData.id} on ${dataName} row ${row_i}`);
  }

  // Write sentence ID
  sheet.getRange(row_i, 1).setValue(sentData.id);
  // Color the header row
  sheet.getRange(row_i, 1, 1, wordPerLine + 1).setBackground("#eee");

  row_i += 1;

  // Write the MWE and CharOffset column headers
  sheet.getRange(row_i, 1).setValue("Span");
  sheet.getRange(row_i, 2).setValue("CharOffset");
  
  // tokens에서 surface, start, end 추출
  var words = sentData.tokens.map((token) => token.surface);
  var starts = sentData.tokens.map((token) => {
    var val = token.start;
    if (val === undefined || val === null) {
      Logger.log("Warning: token.start is missing for token: " + JSON.stringify(token));
      return "";
    }
    return val;
  });
  var ends = sentData.tokens.map((token) => {
    var val = token.end;
    if (val === undefined || val === null) {
      Logger.log("Warning: token.end is missing for token: " + JSON.stringify(token));
      return "";
    }
    return val;
  });
  
  // This sentence is divided into `numRowsForThisSent` rows (bands)
  var numRowsForThisSent = Math.ceil(words.length / wordPerLine);
  
  // Initialize ranges
  var rangesForWords = [];
  var rangesForOffsets = []; // 숨겨진 offset 범위들
  var rangesForCboxes = {}; // This will be mapping from cbox_i (integer from 1 to 9) to ranges (array)
  for (var i = 1; i <= maxMwePerSent; i++) {
    rangesForCboxes[i] = [];
  }
  
  // Prepare ranges for each band
  // 각 band 구조:
  // row_i + headerSectionHeight + i*bandWidth: index 행
  // row_i + headerSectionHeight + i*bandWidth + 1: word 행
  // row_i + headerSectionHeight + i*bandWidth + 2 ~ +10: checkbox 행들 (9개)
  // row_i + headerSectionHeight + i*bandWidth + 11: start 숨김 행
  // row_i + headerSectionHeight + i*bandWidth + 12: end 숨김 행
  // row_i + headerSectionHeight + i*bandWidth + 13: index 숨김 행 (형태소 인덱스)
  for (var i = 0; i < numRowsForThisSent; i++) {
    var bandStartRow = row_i + headerSectionHeight + i * bandWidth;
    
    // Words
    var range = sheet
      .getRange(bandStartRow + 1, colWordStart, 1, wordPerLine)
      .getA1Notation();
    rangesForWords.push(range);
    
    // Offsets (숨김 행에 저장된 start, end 값들을 참조)
    // start와 end를 각각 별도 범위로 저장
    var rangeStart = sheet
      .getRange(bandStartRow + maxMwePerSent + 2, colWordStart, 1, wordPerLine)
      .getA1Notation();
    var rangeEnd = sheet
      .getRange(bandStartRow + maxMwePerSent + 3, colWordStart, 1, wordPerLine)
      .getA1Notation();
    rangesForOffsets.push({ start: rangeStart, end: rangeEnd });
    
    // Checkboxes
    for (var cbox_i = 1; cbox_i <= maxMwePerSent; cbox_i++) {
      var rangeCbox = sheet
        .getRange(bandStartRow + 1 + cbox_i, colWordStart, 1, wordPerLine)
        .getA1Notation();
      rangesForCboxes[cbox_i].push(rangeCbox);
    }
  }
  
  // Fill in cells in the MWE and CharOffset column
  function getFormula(ranges, cbox_i) {
    var subFormulas = [];
    for (var i = 0; i < numRowsForThisSent; i++) {
      var subFormula = `TRIM(JOIN(" ", ARRAYFORMULA(IF(${rangesForCboxes[cbox_i][i]}, ${ranges[i]}, ""))))`;
      subFormulas.push(subFormula);
    }
    var formula = "TRIM(" + subFormulas.join('&" "&') + ")";
    return formula;
  }
  
  // CharOffset 수식 생성 함수
  // 연속된 형태소들을 하나의 범위로 합쳐서 (min_start, max_end) 형태로 출력
  // 각 형태소에서 범위 시작점과 끝점을 찾아 합침
  function getCharOffsetFormula(cbox_i) {
    var bandOffsetFormulas = [];
    
    for (var bandIdx = 0; bandIdx < numRowsForThisSent; bandIdx++) {
      var bandStartRow = row_i + headerSectionHeight + bandIdx * bandWidth;
      var startRow = bandStartRow + maxMwePerSent + 2;  // start 숨김 행
      var endRow = bandStartRow + maxMwePerSent + 3;    // end 숨김 행
      var cbRow = bandStartRow + 1 + cbox_i;            // 체크박스 행
      var colStart = colWordStart;
      
      var segmentParts = [];
      
      // 현재 band에서 실제 형태소 개수 계산
      var wordsInThisBand = Math.min(wordPerLine, words.length - bandIdx * wordPerLine);
      
      // 각 형태소를 순회하며 연속된 그룹 찾기
      for (var offset = 0; offset < wordsInThisBand; offset++) {
        var col = colStart + offset;
        var colLetter = getColumnLetter(col);
        
        var cb = `${colLetter}${cbRow}`;
        var st = `${colLetter}${startRow}`;
        var en = `${colLetter}${endRow}`;
        
        // 현재 형태소가 체크되어 있는지 (실제 형태소가 있는지도 확인)
        var isChecked = `AND(${cb}=TRUE, ISNUMBER(${st}), ISNUMBER(${en}))`;
        
        // 이전 형태소가 체크 안됨 여부 판단
        var prevNotChecked = "";
        if (offset === 0) {
          // 첫 번째 형태소: 이전 band의 마지막 형태소 확인
          if (bandIdx > 0) {
            var prevBandStartRow = row_i + headerSectionHeight + (bandIdx - 1) * bandWidth;
            var prevBandCbRow = prevBandStartRow + 1 + cbox_i;
            var prevBandLastCol = getColumnLetter(colStart + wordPerLine - 1);
            var prevBandLastCb = `${prevBandLastCol}${prevBandCbRow}`;
            prevNotChecked = `OR(${prevBandLastCb}=FALSE, ISBLANK(${prevBandLastCb}))`;
          } else {
            // 첫 번째 band의 첫 번째 형태소: 이전 없음
            prevNotChecked = "TRUE";
          }
        } else {
          // 같은 band 내 이전 형태소 확인
          var prevCol = getColumnLetter(colStart + offset - 1);
          var prevCb = `${prevCol}${cbRow}`;
          prevNotChecked = `OR(${prevCb}=FALSE, ISBLANK(${prevCb}))`;
        }
        
        // 다음 형태소가 체크 안됨 여부 판단
        var nextNotChecked = "";
        if (offset === wordsInThisBand - 1) {
          // 현재 band의 마지막 형태소: 다음 band의 첫 번째 형태소 확인
          if (bandIdx < numRowsForThisSent - 1) {
            var nextBandStartRow = row_i + headerSectionHeight + (bandIdx + 1) * bandWidth;
            var nextBandCbRow = nextBandStartRow + 1 + cbox_i;
            var nextBandFirstCol = getColumnLetter(colWordStart);
            var nextBandFirstCb = `${nextBandFirstCol}${nextBandCbRow}`;
            nextNotChecked = `OR(${nextBandFirstCb}=FALSE, ISBLANK(${nextBandFirstCb}))`;
          } else {
            // 마지막 band의 마지막 형태소: 다음 없음
            nextNotChecked = "TRUE";
          }
        } else {
          // 같은 band 내 다음 형태소 확인
          var nextCol = getColumnLetter(colStart + offset + 1);
          var nextCb = `${nextCol}${cbRow}`;
          nextNotChecked = `OR(${nextCb}=FALSE, ISBLANK(${nextCb}))`;
        }
        
        // 범위 시작점: 체크되어 있고, 이전 형태소가 체크 안됨
        var isRangeStart = `AND(${isChecked}, ${prevNotChecked})`;
        
        // 범위 끝점: 체크되어 있고, 다음 형태소가 체크 안됨
        var isRangeEnd = `AND(${isChecked}, ${nextNotChecked})`;
        
        // 범위 시작점: "(" & start & "," 추가
        var startPart = `IF(${isRangeStart}, "(" & TEXT(${st}, "0") & ",", "")`;
        // 범위 끝점: end & ")" 추가 (콤마 포함)
        var endPart = `IF(${isRangeEnd}, TEXT(${en}, "0") & "),", "")`;
        
        segmentParts.push(startPart);
        segmentParts.push(endPart);
      }
      
      var bandFormula = `TEXTJOIN("", TRUE, ${segmentParts.join(", ")})`;
      bandOffsetFormulas.push(bandFormula);
    }
    
    // 모든 band의 결과를 합치고, 대괄호로 감싸기
    if (bandOffsetFormulas.length === 0) {
      return '""';
    }
    // 모든 band의 결과를 합침
    var combinedResult = `TEXTJOIN("", TRUE, ${bandOffsetFormulas.join(", ")})`;
    // 빈 문자열이면 빈 문자열 반환, 그렇지 않으면 마지막 콤마 제거하고 대괄호로 감싸기
    var finalFormula = `IF(TRIM(${combinedResult})="", "", "[" & REGEXREPLACE(${combinedResult}, ",$", "") & "]")`;
    return finalFormula;
  }
  
  // 열 번호를 문자로 변환 (1 -> A, 2 -> B, ..., 27 -> AA, ...)
  function getColumnLetter(colNum) {
    var result = "";
    while (colNum > 0) {
      var remainder = (colNum - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      colNum = Math.floor((colNum - 1) / 26);
    }
    return result;
  }
  
  for (var cbox_i = 1; cbox_i <= maxMwePerSent; cbox_i++) {
    // Formula for Span (형태소 표면형)
    var formulaSpan = getFormula(rangesForWords, cbox_i);
    sheet.getRange(row_i + cbox_i, 1).setFormula(formulaSpan);
    
    // Formula for CharOffset
    var formulaCharOffset = getCharOffsetFormula(cbox_i);
    sheet.getRange(row_i + cbox_i, 2).setFormula(formulaCharOffset);
  }
  
  // CharOffset 시트의 A열에 통합 CharOffset 수식 추가 (모든 MWE 행의 B열 CharOffset을 합침)
  var charOffsetRangeStart = row_i + 1; // 첫 번째 MWE 행
  var charOffsetRangeEnd = row_i + maxMwePerSent; // 마지막 MWE 행
  var colB = 2; // B열은 2번째 열
  var colBLetter = getColumnLetter(colB);
  var sheetNameForFormula = sheet.getName(); // 시트 이름을 수식에서 사용
  
  // 모든 MWE 행의 B열 CharOffset을 공백으로 연결하는 수식
  // 다른 시트를 참조하므로 시트 이름 포함
  // 형태소를 체크하지 않은 경우 빈 문자열을 반환하고, 그 경우 셀도 비워지도록 처리
  // Google Sheets는 작은따옴표를 자동으로 제거하므로 작은따옴표 없이 사용
  var joinedResult = `TEXTJOIN(" ", TRUE, ${sheetNameForFormula}!${colBLetter}${charOffsetRangeStart}:${colBLetter}${charOffsetRangeEnd})`;
  var combinedCharOffsetFormula = `IF(TRIM(${joinedResult})="", "", ${joinedResult})`;
  
  // 전역 문장 번호 행의 CharOffset 시트 A열에 수식 추가 (첫 번째 문장은 1행, 두 번째 문장은 2행...)
  // 주의: globalSentenceNumber가 null이 아니고 0보다 큰 경우에만 설정
  if (globalSentenceNumber !== null && globalSentenceNumber > 0 && charOffsetSheet) {
    // 기존 유효성 검사 제거 (에러 방지)
    try {
      charOffsetSheet.getRange(globalSentenceNumber, 1).clearDataValidations();
    } catch (e) {
      // 유효성 검사가 없어도 무시
    }
    
    // 수식 설정
    try {
      charOffsetSheet.getRange(globalSentenceNumber, 1).setFormula(combinedCharOffsetFormula);
    } catch (e) {
      Logger.log(`[에러] CharOffset 수식 설정 실패 (문장 ID: ${sentData.id}, 전역 번호: ${globalSentenceNumber}): ${e}`);
    }
  }

  // Write a border encompassing the range to be filled in by annotators
  sheet
    .getRange(row_i + 1, 1, maxMwePerSent, 2)
    .setBorder(true, true, true, true, false, false);

  // Write the "Sentence" column with context
  var sentenceRow = row_i + maxMwePerSent + 1;
  sheet.getRange(sentenceRow, 1).setValue("Sentence");
  
  // target_sentence 표시 (기존 text 필드도 지원)
  var targetText = sentData.target_sentence || sentData.text || "";
  sheet.getRange(sentenceRow, 2).setValue(targetText);
  
  // context_left 표시 (회색 텍스트)
  if (sentData.context_left) {
    var contextLeftRow = sentenceRow + 1;
    sheet.getRange(contextLeftRow, 1).setValue("Context (Left)");
    sheet.getRange(contextLeftRow, 2).setValue(sentData.context_left);
    sheet.getRange(contextLeftRow, 1, 1, 2).setBackground("#f5f5f5").setFontColor("#999999");
  }
  
  // context_right 표시 (회색 텍스트)
  if (sentData.context_right) {
    var contextRightRow = sentenceRow + (sentData.context_left ? 2 : 1);
    sheet.getRange(contextRightRow, 1).setValue("Context (Right)");
    sheet.getRange(contextRightRow, 2).setValue(sentData.context_right);
    sheet.getRange(contextRightRow, 1, 1, 2).setBackground("#f5f5f5").setFontColor("#999999");
  }

  // ==========
  // 형태소 배치 섹션 작성

  // 첫 번째 band 시작 행
  var firstBandRow = row_i + headerSectionHeight;

  // Write the Span column header for each band
  function writeSpanColumnHeader(bandRow) {
    sheet.getRange(bandRow + 1, 1).setValue("Span");
  }

  // Write words, indices, checkboxes, and hidden offset rows
  // 성능 최적화: band별로 배치 처리
  var bandStartRow = firstBandRow;
  
  // 첫 번째 band 헤더 작성
  writeSpanColumnHeader(firstBandRow);
  
  for (var bandIdx = 0; bandIdx < numRowsForThisSent; bandIdx++) {
    bandStartRow = firstBandRow + bandIdx * bandWidth;
    var wordsInThisBand = Math.min(wordPerLine, words.length - bandIdx * wordPerLine);
    var bandStartWordIdx = bandIdx * wordPerLine;
    
    // 배치 데이터 준비
    var batchIndices = [];
    var batchWords = [];
    var batchWordsFormula = []; // 수식이 필요한 경우
    var batchStarts = [];
    var batchEnds = [];
    var batchIndicesHidden = [];
    
    for (var i = 0; i < wordsInThisBand; i++) {
      var word_i = bandStartWordIdx + i;
      batchIndices.push(word_i + 1);
      
      // 단어 처리
      if (words[word_i].startsWith("'") || words[word_i].includes("/")) {
        batchWords.push(null); // 나중에 수식으로 설정
        batchWordsFormula.push({idx: i, formula: `="${words[word_i]}"`});
      } else {
        batchWords.push(words[word_i]);
      }
      
      // offset 값 처리
      var startVal = starts[word_i];
      var endVal = ends[word_i];
      if (typeof startVal !== 'number') {
        startVal = Number(startVal);
        if (isNaN(startVal)) startVal = "";
      }
      if (typeof endVal !== 'number') {
        endVal = Number(endVal);
        if (isNaN(endVal)) endVal = "";
      }
      batchStarts.push(startVal);
      batchEnds.push(endVal);
      batchIndicesHidden.push(word_i + 1);
    }
    
    // 배치로 한 번에 설정 (API 호출 최소화)
    sheet.getRange(bandStartRow, colWordStart, 1, wordsInThisBand).setValues([batchIndices]);
    sheet.getRange(bandStartRow + 1, colWordStart, 1, wordsInThisBand).setValues([batchWords.map(w => w !== null ? w : "")]);
    sheet.getRange(bandStartRow + maxMwePerSent + 2, colWordStart, 1, wordsInThisBand).setValues([batchStarts]);
    sheet.getRange(bandStartRow + maxMwePerSent + 3, colWordStart, 1, wordsInThisBand).setValues([batchEnds]);
    sheet.getRange(bandStartRow + maxMwePerSent + 4, colWordStart, 1, wordsInThisBand).setValues([batchIndicesHidden]);
    
    // 수식이 필요한 단어 개별 설정
    for (var f = 0; f < batchWordsFormula.length; f++) {
      var formulaInfo = batchWordsFormula[f];
      sheet.getRange(bandStartRow + 1, colWordStart + formulaInfo.idx).setFormula(formulaInfo.formula);
    }
    
    // 체크박스 배치 적용
    sheet.getRange(bandStartRow + 2, colWordStart, maxMwePerSent, wordsInThisBand).insertCheckboxes();
    
    // 스타일 배치 적용 (숨김 행들)
    var styleRange = sheet.getRange(bandStartRow + maxMwePerSent + 2, colWordStart, 3, wordsInThisBand);
    var bgColors = [];
    var fontColors = [];
    for (var r = 0; r < 3; r++) {
      var bgRow = [];
      var fontRow = [];
      for (var c = 0; c < wordsInThisBand; c++) {
        bgRow.push("#f0f0f0");
        fontRow.push("#cccccc");
      }
      bgColors.push(bgRow);
      fontColors.push(fontRow);
    }
    styleRange.setBackgrounds(bgColors);
    styleRange.setFontColors(fontColors);
    styleRange.setFontSize(8);
    
    // 다음 band를 위한 헤더 작성
    if (bandIdx < numRowsForThisSent - 1) {
      writeSpanColumnHeader(bandStartRow + bandWidth);
    }
  }
  
  // 마지막 band 이후 행 높이 조정 (시각적 구분)
  var lastBandEndRow = bandStartRow + bandWidth;
  sheet.insertRowAfter(lastBandEndRow);

  setNextWriteRowPointer(ss, dataName, rowStart + estimateRowsNeededForSentence(sentData));
}

/********************
 * Utilities (Internal)
 ********************/

function createSpreadsheet(folder) {
  // Create a spreadsheet and move it into the output folder.
  var spreadsheet = SpreadsheetApp.create("New Spreadsheet");
  var file = DriveApp.getFileById(spreadsheet.getId());
  file.moveTo(folder);
  return spreadsheet;
}
  
function cleanFolder(folder) {
  // Move all existing output files in the target folder to trash.
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
  }
}
  
function findChildFolderByName(projectFolder, targetFolderName) {
  var subFolders = projectFolder.getFolders();
  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    if (folder.getName() === targetFolderName) {
      return folder;
    }
  }

  throw new Error(
    `Required folder "${targetFolderName}" was not found inside "${projectFolder.getName()}". Check projectFolderId and folder names in User Settings.`
  );
}
  
function findFileByName(folder, targetFileName) {
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName() == targetFileName) {
      return file;
    }
  }

  throw new Error(`Required file "${targetFileName}" was not found inside "${folder.getName()}".`);
}
  
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
