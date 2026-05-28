/**
 * 病棟ラウンド集計アプリ — Google Apps Script
 *
 * 【デプロイ手順】
 * 1. Googleスプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. デプロイ → 新しいデプロイ → 種類:「ウェブアプリ」
 *    - 実行ユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイ後に表示されるURLを index.html の設定画面に入力
 *
 * 【集計シートの更新方法】
 * スプレッドシートのメニュー「ラウンド集計」→「集計シートを更新」を押す
 *
 * 【介入開始日の設定方法】
 * スプレッドシートのメニュー「ラウンド集計」→「介入開始日を設定」で日付を入力
 * 設定後に「集計シートを更新」を実行すると介入フラグ列が更新される
 */

const SHEET_NAME = 'ラウンド記録';

const HEADERS = [
  '送信日時',           // A
  '日付',               // B
  'チーム',             // C
  'ラウンド種別',        // D
  '救急外来専属医師',    // E
  'HCU専属医師',        // F
  '兼任フラグ',          // G
  '開催フラグ',          // H
  '開始時刻',            // I
  '終了時刻',            // J
  '参加人数',            // K
  '途中退席人数',        // L
  '不参加人数',          // M
  '参加者詳細(JSON)',    // N
  '記録信頼性',          // O
  '除外理由',            // P
];

// ── POST受信 ──────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // メンバー設定の保存
    if (data.action === 'saveMembers') {
      PropertiesService.getScriptProperties()
        .setProperty(`members_${data.team}`, JSON.stringify(data.members || []));
      return jsonResponse({ status: 'ok' });
    }

    const sheet = getOrCreateSheet();
    const row = [
      new Date(),
      data.date        || '',
      data.team        || '',
      data.roundType   || '',
      data.erDoctor    || 'いない',
      data.hcuDoctor   || 'いない',
      data.concurrent  || '該当なし',
      data.held ? '開催あり' : '開催なし',
      data.startTime   || '',
      data.endTime     || '',
      data.participantCount ?? 0,
      data.partialCount     ?? 0,
      data.absentCount      ?? 0,
      JSON.stringify(data.members || []),
      data.valid ? '含める' : '含めない',
      data.exclusionReason || '',
    ];
    sheet.appendRow(row);
    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── GET（動作確認 & メンバー取得 & 集計） ──────────
function doGet(e) {
  if (e && e.parameter) {
    if (e.parameter.action === 'getMembers') {
      const props = PropertiesService.getScriptProperties();
      return jsonResponse({
        status: 'ok',
        A: JSON.parse(props.getProperty('members_A') || '[]'),
        B: JSON.parse(props.getProperty('members_B') || '[]'),
      });
    }
    if (e.parameter.action === 'getStats') {
      return getWeekdayStats();
    }
  }
  return ContentService
    .createTextOutput('病棟ラウンド集計 API は正常に動作しています。')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getWeekdayStats() {
  const src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!src) return jsonResponse({ status: 'error', message: 'シートが見つかりません' });

  const rows = src.getDataRange().getValues();
  // dow 1=月 2=火 3=水 4=木 5=金  値: { dateStr: count }
  const byDow = { 1:{}, 2:{}, 3:{}, 4:{}, 5:{} };

  for (let r = 1; r < rows.length; r++) {
    const dateVal = rows[r][1];   // B 日付
    const valid   = rows[r][14];  // O 記録信頼性
    if (!dateVal || valid !== '含める') continue;

    const d   = new Date(dateVal);
    const dow = d.getDay();
    if (dow < 1 || dow > 5) continue;

    const dateStr = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    byDow[dow][dateStr] = (byDow[dow][dateStr] || 0) + 1;
  }

  const DAY_NAMES = { 1:'月', 2:'火', 3:'水', 4:'木', 5:'金' };
  const stats = [1, 2, 3, 4, 5].map(dow => {
    const entries      = Object.values(byDow[dow]);
    const dayCount     = entries.length;
    const totalRecords = entries.reduce((s, n) => s + n, 0);
    const avgPerDay    = dayCount > 0 ? Math.round(totalRecords / dayCount * 10) / 10 : 0;
    const daysOnTarget = entries.filter(n => n >= 2).length;
    return { dow, dayName: DAY_NAMES[dow], dayCount, totalRecords, avgPerDay, daysOnTarget };
  });

  return jsonResponse({ status: 'ok', stats });
}

// ══════════════════════════════════════════════════
//  カスタムメニュー
// ══════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ラウンド集計')
    .addItem('集計シートを更新', 'flattenAll')
    .addSeparator()
    .addItem('介入開始日を設定', 'setInterventionDate')
    .addItem('介入開始日を確認', 'showInterventionDate')
    .addToUi();
}

// ── 介入開始日の設定・確認 ────────────────────────
function setInterventionDate() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties()
                    .getProperty('INTERVENTION_DATE') || '未設定';
  const res = ui.prompt(
    '介入開始日の設定',
    `現在の設定: ${current}\n\n` +
    '新しい介入開始日を入力してください（例: 2026-07-01）\n' +
    '※ この日以降のデータが「介入後」になります',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const val = res.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(val)) {
    ui.alert('形式が正しくありません。YYYY-MM-DD 形式で入力してください（例: 2026-07-01）');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('INTERVENTION_DATE', val);
  ui.alert(`介入開始日を ${val} に設定しました。\n「集計シートを更新」を実行してください。`);
}

function showInterventionDate() {
  const val = PropertiesService.getScriptProperties()
                .getProperty('INTERVENTION_DATE') || '未設定';
  SpreadsheetApp.getUi().alert(`現在の介入開始日: ${val}`);
}

// ══════════════════════════════════════════════════
//  集計シート生成
// ══════════════════════════════════════════════════

function flattenAll() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SHEET_NAME);
  if (!src) {
    SpreadsheetApp.getUi().alert('「ラウンド記録」シートが見つかりません');
    return;
  }
  const rows = src.getDataRange().getValues();
  if (rows.length <= 1) {
    SpreadsheetApp.getUi().alert('記録データがありません');
    return;
  }

  // 介入開始日（未設定なら空文字）
  const interventionDate = PropertiesService.getScriptProperties()
                             .getProperty('INTERVENTION_DATE') || '';

  const memberRows = [];
  const depRows    = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];

    const sentAt     = row[0];   // A 送信日時
    const date       = row[1];   // B 日付
    const team       = row[2];   // C チーム
    const roundType  = row[3];   // D ラウンド種別
    const erDoc      = row[4];   // E 救急外来専属医師
    const hcuDoc     = row[5];   // F HCU専属医師
    const concurrent = row[6];   // G 兼任フラグ
    const held       = row[7];   // H 開催フラグ
    const startTime  = row[8];   // I 開始時刻
    const endTime    = row[9];   // J 終了時刻
    const valid      = row[14];  // O 記録信頼性
    const jsonStr    = row[13];  // N 参加者詳細JSON

    // 介入フラグ（介入開始日未設定なら「未設定」）
    const flag = !interventionDate
      ? '未設定'
      : (String(date) >= interventionDate ? '介入後' : '介入前');

    let members = [];
    try { members = JSON.parse(jsonStr); } catch (e) { continue; }

    members.forEach(m => {
      const status = m.status     || '';
      const name   = m.name       || '';
      const role   = m.role       || '';
      const deps   = m.departures || [];

      // 参加区分
      let category;
      switch (status) {
        case '参加':     category = '完全参加'; break;
        case '途中退席': category = '部分参加'; break;
        case '不参加':   category = '不在';     break;
        default:         category = '対象外';   break;
      }

      // 不参加理由の整形
      const absReason = m.absenceReason === 'その他' && m.absenceReasonOther
        ? `その他（${m.absenceReasonOther}）`
        : (m.absenceReason || '');

      // メンバー集計シートへの1行
      memberRows.push([
        date,                          // A 日付
        flag,                          // B 介入フラグ
        team,                          // C チーム
        roundType,                     // D ラウンド種別
        held,                          // E 開催フラグ
        erDoc,                         // F 救急外来専属医師
        hcuDoc,                        // G HCU専属医師
        concurrent,                    // H 兼任フラグ
        startTime,                     // I 開始時刻
        endTime,                       // J 終了時刻
        name,                          // K 名前
        role,                          // L 職種
        status,                        // M ステータス
        category,                      // N 参加区分
        deps.length,                   // O 退席回数
        absReason,                     // P 不参加理由
        m.absenceReturnTime || '',     // Q 不参加後復帰時刻
        valid,                         // R 記録信頼性
        sentAt,                        // S 送信日時
      ]);

      // 退席詳細シートへの行（途中退席・不参加どちらも対象）
      deps.forEach((d, j) => {
        const reason = d.reason === 'その他' && d.reasonOther
          ? `その他（${d.reasonOther}）`
          : (d.reason || '');
        depRows.push([
          date,                          // A 日付
          flag,                          // B 介入フラグ
          team,                          // C チーム
          roundType,                     // D ラウンド種別
          erDoc,                         // E 救急外来専属医師
          hcuDoc,                        // F HCU専属医師
          name,                          // G 名前
          role,                          // H 職種
          status,                        // I ステータス
          j + 1,                         // J 退席番号
          d.departureTime || '',         // K 退席時刻
          reason,                        // L 退席理由
          d.returnTime    || '',         // M 復帰時刻
          d.returnTime ? 'あり' : 'なし', // N 復帰
          valid,                         // O 記録信頼性
          sentAt,                        // P 送信日時
        ]);
      });
    });
  }

  writeSheet(ss, 'メンバー集計', [
    '日付','介入フラグ','チーム','ラウンド種別','開催フラグ',
    '救急外来専属医師','HCU専属医師','兼任フラグ',
    '開始時刻','終了時刻',
    '名前','職種','ステータス','参加区分',
    '退席回数','不参加理由','不参加後復帰時刻',
    '記録信頼性','送信日時',
  ], memberRows, '#0f9d58');

  writeSheet(ss, '退席詳細', [
    '日付','介入フラグ','チーム','ラウンド種別',
    '救急外来専属医師','HCU専属医師',
    '名前','職種','ステータス',
    '退席番号','退席時刻','退席理由','復帰時刻','復帰',
    '記録信頼性','送信日時',
  ], depRows, '#e67c00');

  SpreadsheetApp.getUi().alert(
    `集計シートを更新しました ✓\n` +
    `介入開始日: ${interventionDate || '未設定'}\n` +
    `メンバー集計: ${memberRows.length} 行\n` +
    `退席詳細: ${depRows.length} 行`
  );
}

// ── シート書き込みヘルパー ────────────────────────
function writeSheet(ss, name, headers, rows, color) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(name);
  }
  sheet.appendRow(headers);
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground(color);
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.autoResizeColumns(1, headers.length);
}

// ══════════════════════════════════════════════════
//  ヘルパー
// ══════════════════════════════════════════════════

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(9, 80);
    sheet.setColumnWidth(10, 80);
    sheet.setColumnWidth(14, 200);
    sheet.setColumnWidth(16, 200);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
