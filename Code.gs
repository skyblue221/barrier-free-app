const DRIVE_FOLDER_ID = '1Zslt_Mkwrj5CwOjL7_hdB7q8ZolmhRTG';
const MANAGER_EMAIL = 'smlee@jpurme.org';

// ── 시트 초기 세팅 ────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 신청 시트
  let apply = ss.getSheetByName('신청');
  if (!apply) apply = ss.insertSheet('신청');
  apply.getRange(1, 1, 1, 8).setValues([[
    '타임스탬프', '신청자', '상호명', '주소', '신청항목', '연락처', '메모', '상태'
  ]]);
  apply.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#E8F5E9');
  apply.setColumnWidth(1, 160);
  apply.setColumnWidth(4, 200);
  apply.setColumnWidth(7, 200);

  // 모니터링 시트
  let monitor = ss.getSheetByName('모니터링');
  if (!monitor) monitor = ss.insertSheet('모니터링');
  monitor.getRange(1, 1, 1, 14).setValues([[
    '타임스탬프', '장소명', '주소', '모니터링날짜',
    '점자메뉴판', '경사로·출입접근성', '화장실접근성', '안내판가독성',
    '조치필요여부', '종합의견', '추가확인항목', '기록자', '상태', '사진링크'
  ]]);
  monitor.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#E3F2FD');
  monitor.setColumnWidth(1, 160);
  monitor.setColumnWidth(3, 200);
  monitor.setColumnWidth(10, 250);
  monitor.setColumnWidth(14, 300);

  // 통합 시트
  let unified = ss.getSheetByName('통합현황');
  if (!unified) unified = ss.insertSheet('통합현황');
  unified.getRange(1, 1, 1, 10).setValues([[
    '타임스탬프', '유형', '장소·상호명', '주소', '신청·점검항목', '조치필요여부', '종합내용', '담당자·기록자', '상태', '사진링크'
  ]]);
  unified.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#FFF8E1');
  unified.setColumnWidth(1, 160);
  unified.setColumnWidth(4, 200);
  unified.setColumnWidth(7, 250);

  // ── 상점 추천 시트 (신규 추가) ──────────────────
  let recommend = ss.getSheetByName('상점추천');
  if (!recommend) recommend = ss.insertSheet('상점추천');
  recommend.getRange(1, 1, 1, 8).setValues([[
    '타임스탬프', '상호명', '주소', '좋은점', '추천인', '연락처', '한줄소개', '상태'
  ]]);
  recommend.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#FCE4EC');
  recommend.setColumnWidth(1, 160);
  recommend.setColumnWidth(2, 150);
  recommend.setColumnWidth(3, 200);
  recommend.setColumnWidth(4, 250);
  recommend.setColumnWidth(5, 100);
  recommend.setColumnWidth(6, 130);
  recommend.setColumnWidth(7, 250);
  recommend.setColumnWidth(8, 100);

  // ── 활동 일지 시트 (신규 추가) ──────────────────
  let journal = ss.getSheetByName('활동일지');
  if (!journal) journal = ss.insertSheet('활동일지');
  journal.getRange(1, 1, 1, 8).setValues([[
    '타임스탬프', '활동날짜', '활동시간', '팀', '팀원(작성자)', '방문상점', '활동소감', '사진링크'
  ]]);
  journal.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#F3E8FF');
  journal.setColumnWidth(1, 160);
  journal.setColumnWidth(2, 120);
  journal.setColumnWidth(3, 90);
  journal.setColumnWidth(4, 80);
  journal.setColumnWidth(5, 150);
  journal.setColumnWidth(6, 300);
  journal.setColumnWidth(7, 350);
  journal.setColumnWidth(8, 400);

  // 배리어프리 마을 만들기 탭 삭제
  const old = ss.getSheetByName('배리어프리 마을 만들기');
  if (old) ss.deleteSheet(old);

  // 시트1 삭제
  const sheet1 = ss.getSheetByName('시트1');
  if (sheet1) ss.deleteSheet(sheet1);

  SpreadsheetApp.getUi().alert('✅ 시트 세팅 완료!');
}

// ── 통합현황 갱신 ─────────────────────────────────
function refreshUnified() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unified = ss.getSheetByName('통합현황');
  const applySheet = ss.getSheetByName('신청');
  const monitorSheet = ss.getSheetByName('모니터링');

  // 기존 데이터 삭제 (헤더 제외)
  const lastRow = unified.getLastRow();
  if (lastRow > 1) unified.deleteRows(2, lastRow - 1);

  const rows = [];

  // 신청 데이터
  const applyData = applySheet.getDataRange().getValues();
  for (let i = 1; i < applyData.length; i++) {
    const r = applyData[i];
    if (!r[0]) continue;
    rows.push([
      r[0],           // 타임스탬프
      '보급 신청',     // 유형
      r[2],           // 상호명
      r[3],           // 주소
      r[4],           // 신청항목
      '',             // 조치필요여부
      r[6] || '',     // 메모
      r[1] || '',     // 신청자
      r[7] || '신청 접수', // 상태
      ''              // 사진링크
    ]);
  }

  // 모니터링 데이터
  const monitorData = monitorSheet.getDataRange().getValues();
  for (let i = 1; i < monitorData.length; i++) {
    const r = monitorData[i];
    if (!r[0]) continue;
    const ratings = [
      r[4] ? '점자메뉴판:' + r[4] : '',
      r[5] ? '경사로:' + r[5] : '',
      r[6] ? '화장실:' + r[6] : '',
      r[7] ? '안내판:' + r[7] : ''
    ].filter(Boolean).join(' / ');
    rows.push([
      r[0],           // 타임스탬프
      '모니터링',      // 유형
      r[1],           // 장소명
      r[2],           // 주소
      ratings,        // 항목별 평가
      r[8] || '',     // 조치필요여부
      r[9] || '',     // 종합의견
      r[11] || '',    // 기록자
      r[12] || '확인 중', // 상태
      r[13] || ''     // 사진링크
    ]);
  }

  // 상점 추천 데이터
  const recommendSheet = ss.getSheetByName('상점추천');
  if (recommendSheet) {
    const recommendData = recommendSheet.getDataRange().getValues();
    for (let i = 1; i < recommendData.length; i++) {
      const r = recommendData[i];
      if (!r[0]) continue;
      rows.push([
        r[0],       // 타임스탬프
        '상점 추천', // 유형
        r[1],       // 상호명
        r[2],       // 주소
        r[3],       // 좋은 점
        '',         // 조치필요여부
        r[6] || '', // 한줄소개
        r[4] || '', // 추천인
        r[7] || '추천', // 상태
        ''          // 사진링크
      ]);
    }
  }

  // 날짜 최신순 정렬
  rows.sort((a, b) => new Date(b[0]) - new Date(a[0]));

  if (rows.length > 0) {
    unified.getRange(2, 1, rows.length, 10).setValues(rows);

    // 유형별 색상
    for (let i = 0; i < rows.length; i++) {
      const color = rows[i][1] === '보급 신청' ? '#E8F5E9' : '#E3F2FD';
      unified.getRange(i + 2, 1, 1, 10).setBackground(color);
    }
  }

  SpreadsheetApp.getUi().alert(`✅ 통합현황 갱신 완료! 총 ${rows.length}건`);
}

// ── 메뉴 추가 ─────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('배리어프리 관리')
    .addItem('통합현황 갱신', 'refreshUnified')
    .addItem('시트 초기 세팅', 'setup')
    .addToUi();
}

// ── 데이터 수신 (doPost) ──────────────────────────
function doPost(e) {
  try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  Logger.log('doPost type: ' + data.type);

  if (data.type === 'apply') {
    const sheet = ss.getSheetByName('신청');
    sheet.appendRow([
      data.timestamp,
      data.applyname || '',
      data.name,
      data.addr,
      data.items,
      data.phone,
      data.memo || '',
      '신청 접수',
      data.privacy || 'N'
    ]);
    MailApp.sendEmail(
      MANAGER_EMAIL,
      `[배리어프리 신청] ${data.name} — ${data.items}`,
      `새로운 신청이 접수되었습니다.\n\n` +
      `■ 신청자: ${data.applyname || '미입력'}\n` +
      `■ 상호명: ${data.name}\n` +
      `■ 주소: ${data.addr}\n` +
      `■ 신청 항목: ${data.items}\n` +
      `■ 연락처: ${data.phone}\n` +
      `■ 메모: ${data.memo || '없음'}\n` +
      `■ 접수 시간: ${data.timestamp}`
    );

  } else if (data.type === 'monitor') {
    Logger.log('monitor 수신 - 장소: ' + data.place + ', 사진: ' + (data.photos ? data.photos.length : 0) + '장');
    let photoLinks = '';
    if (data.photos && data.photos.length > 0) {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const subFolder = folder.createFolder(data.place + '_' + data.date);
      const links = data.photos.map((photo, idx) => {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(photo.data),
          photo.mimeType,
          `${data.place}_${idx + 1}.jpg`
        );
        const file = subFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
      });
      photoLinks = links.join('\n');
    }
    const sheet = ss.getSheetByName('모니터링');
    sheet.appendRow([
      data.timestamp, data.place, data.addr, data.date,
      data.r1 || '', data.r2 || '', data.r3 || '', data.r4 || '',
      data.action || '', data.comment,
      data.extras || '', data.author || '',
      '확인 중', photoLinks
    ]);

  } else if (data.type === 'recommend') {
    // ── 상점 추천 저장 ──────────────────────────
    const sheet = ss.getSheetByName('상점추천');
    sheet.appendRow([
      data.timestamp,
      data.name || '',
      data.addr || '',
      data.items || '',
      data.recommender || '',
      data.phone || '',
      data.memo || '',
      '추천'
    ]);
    MailApp.sendEmail(
      MANAGER_EMAIL,
      `[배리어프리 추천] ${data.name} — ${data.items}`,
      `새로운 상점 추천이 접수되었습니다.\n\n` +
      `■ 상호명: ${data.name}\n` +
      `■ 주소: ${data.addr}\n` +
      `■ 좋은 점: ${data.items}\n` +
      `■ 추천인: ${data.recommender}\n` +
      `■ 연락처: ${data.phone || '미입력'}\n` +
      `■ 한줄소개: ${data.memo || '없음'}\n` +
      `■ 접수 시간: ${data.timestamp}`
    );

  } else if (data.type === 'journal') {
    // ── 활동 일지 저장 + 사진링크 자동 수집 ────────
    const sheet = ss.getSheetByName('활동일지');
    const monitorSheet = ss.getSheetByName('모니터링');

    // 같은 날짜 + 팀원 이름 매칭으로 사진링크 자동 수집
    const photoLinks = [];
    if (monitorSheet && data.date && data.author) {
      const monitorRows = monitorSheet.getDataRange().getValues();
      const authorNames = data.author.split(',').map(s => s.trim());
      for (let i = 1; i < monitorRows.length; i++) {
        const r = monitorRows[i];
        if (!r[0]) continue;
        // 날짜 일치 (r[3]: 모니터링날짜)
        const rowDate = String(r[3] || '').trim();
        if (rowDate !== data.date) continue;
        // 기록자 이름이 팀원 중 하나와 일치
        const rowAuthor = String(r[11] || '').trim();
        const matched = authorNames.some(name => rowAuthor.includes(name));
        if (!matched) continue;
        // 사진링크(r[13]) 수집
        const links = String(r[13] || '').trim();
        if (links) {
          const storeName = String(r[1] || '');
          // 상점명과 함께 표시
          links.split('\n').filter(Boolean).forEach(url => {
            photoLinks.push(storeName + ': ' + url);
          });
        }
      }
    }

    sheet.appendRow([
      data.timestamp,              // 타임스탬프
      data.date || '',             // 활동날짜
      data.time || '',             // 활동시간
      data.team || '',             // 팀
      data.author || '',           // 팀원(작성자)
      data.stores || '',           // 방문상점 (쉼표 구분)
      data.memo || '',             // 활동소감
      photoLinks.join('\n') || ''  // 사진링크 (상점명: URL 형식)
    ]);

  } else if (data.type === 'uploadPhoto') {
    // ── 사진 단건 업로드 → Drive 저장 → URL 반환 ──
    Logger.log('사진 업로드: ' + data.place + ' ' + (data.photoIndex + 1) + '번째');
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const folderName = (data.place || 'unknown') + '_' + (data.date || '');
    let subFolder;
    const existFolders = folder.getFoldersByName(folderName);
    subFolder = existFolders.hasNext() ? existFolders.next() : folder.createFolder(folderName);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.photo),
      'image/jpeg',
      (data.place || 'photo') + '_' + (data.photoIndex + 1) + '.jpg'
    );
    const file = subFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const photoUrl = file.getUrl();
    Logger.log('사진 저장 완료: ' + photoUrl);
    return makeResponse({ result: 'ok', url: photoUrl });

  } else if (data.type === 'addPhotoLink') {
    // ── 사진링크 시트에 추가 ─────────────────────
    const monSheet = ss.getSheetByName('모니터링');
    const monRows = monSheet.getDataRange().getValues();
    for (let i = 1; i < monRows.length; i++) {
      if (String(monRows[i][0]) === String(data.timestamp)) {
        const existing = String(monRows[i][13] || '').trim();
        monSheet.getRange(i + 1, 14).setValue(existing ? existing + '
' + data.url : data.url);
        Logger.log('사진링크 추가: ' + data.url);
        break;
      }
    }
    return makeResponse({ result: 'ok' });

  } else if (data.type === 'updateStatus') {
    const sheetName = data.recordType === 'apply' ? '신청' : '모니터링';
    const sheet = ss.getSheetByName(sheetName);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.timestamp)) {
        const statusCol = data.recordType === 'apply' ? 8 : 13;
        sheet.getRange(i + 1, statusCol).setValue(data.status);
        break;
      }
    }
  }

  return makeResponse({ result: 'ok' });
  } catch(err) {
    Logger.log('doPost 오류: ' + err.toString());
    return makeResponse({ result: 'error', message: err.toString() });
  }
}

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 사진 업로드 전용 엔드포인트 ──────────────────────
function doUploadPhoto(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const folderName = (data.place || 'unknown') + '_' + (data.date || '');
    let subFolder;
    const folders = folder.getFoldersByName(folderName);
    subFolder = folders.hasNext() ? folders.next() : folder.createFolder(folderName);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.photo.data),
      'image/jpeg',
      (data.place || 'photo') + '_' + (data.photoIndex + 1) + '.jpg'
    );
    const file = subFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return makeResponse({ result: 'ok', url: file.getUrl() });
  } catch(err) {
    Logger.log('사진 업로드 오류: ' + err.toString());
    return makeResponse({ result: 'error', message: err.toString() });
  }
}

// ── 데이터 조회 (doGet) ───────────────────────────
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const type = e.parameter.type;

  if (type === 'apply') {
    const sheet = ss.getSheetByName('신청');
    const rows = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        type: 'apply',
        timestamp: r[0],
        applyname: r[1] || '',
        name: r[2],
        addr: r[3],
        items: r[4],
        phone: r[5],
        memo: r[6] || '',
        action: r[7] || '신청 접수'
      });
    }
    return makeResponse({ result: 'ok', records });

  } else if (type === 'monitor') {
    const sheet = ss.getSheetByName('모니터링');
    const rows = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        type: 'monitor',
        timestamp: r[0],
        name: r[1],
        addr: r[2],
        date: r[3],
        r1: r[4] || '',
        r2: r[5] || '',
        r3: r[6] || '',
        r4: r[7] || '',
        action: r[8] || '',
        comment: r[9] || '',
        extras: r[10] || '',
        author: r[11] || '',
        status: r[12] || '',
        photoLinks: r[13] || ''
      });
    }
    return makeResponse({ result: 'ok', records });

  } else if (type === 'recommend') {
    // ── 상점 추천 조회 ──────────────────────────
    const sheet = ss.getSheetByName('상점추천');
    if (!sheet) {
      return makeResponse({ result: 'ok', records: [] });
    }
    const rows = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        type: 'recommend',
        timestamp: String(r[0]),
        name: String(r[1] || ''),
        addr: String(r[2] || ''),
        items: String(r[3] || ''),
        recommender: String(r[4] || ''),
        phone: String(r[5] || ''),
        memo: String(r[6] || ''),
        action: String(r[7] || '추천')
      });
    }
    return makeResponse({ result: 'ok', records });

  } else if (type === 'journal') {
    // ── 활동 일지 조회 ──────────────────────────
    const sheet = ss.getSheetByName('활동일지');
    if (!sheet) {
      return makeResponse({ result: 'ok', records: [] });
    }
    const rows = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        timestamp: String(r[0]),
        date: String(r[1] || ''),
        time: String(r[2] || ''),
        team: String(r[3] || ''),
        author: String(r[4] || ''),
        stores: String(r[5] || ''),
        memo: String(r[6] || ''),
        photoLinks: String(r[7] || '')
      });
    }
    return makeResponse({ result: 'ok', records });
  }

  return makeResponse({ result: 'error' });
}
