/**
 * ============================================================
 * JWNET マニフェスト登録　収集運搬業者さん用 入力フォーム ＋ CSV展開シート
 * 版: v2.1（2026-10-13）　※v2.1：必須項目（引渡し担当者・数量の確定者・登録区分）も、任意の質問として残した
 * ============================================================
 *
 * 【この仕組みの前提】
 *  FAXをやめて、産廃（収集運搬）業者さんに、Googleフォームへ直接入力してもらう。
 *  入力された内容は、野村工務店さん側で確認（✔）してから、JWNET用CSVに展開される。
 *
 * 【できあがるもの】
 *  1枚目「フォーム回答」 … 業者さんの入力が溜まる。右端の「確認」列に ✔ を付けると出力対象になる
 *  2枚目「JWNET出力」    … ✔ 済みの回答を、関数で JWNET「新規登録（読込）」CSVの75列に自動展開
 *  3枚目「設定」         … 現場・業者・コードの一覧と固定値（直すとフォームの選択肢に反映）
 *  4枚目「出力管理」     … 確認待ち件数／出力前のチェック結果
 *
 * 【使い方（初回のみ）】
 *  1. 新しいスプレッドシートを作り「拡張機能 > Apps Script」を開く
 *  2. このコードを全部貼り付けて保存
 *  3. 関数「setupAll」を選んで「実行」（初回は承認が必要）
 *  4. 「設定」シートの「（仮）」を本物に書き換える
 *  5. メニュー「JWNET > フォームの選択肢を更新」
 *  6. メニュー「JWNET > 入力フォームのURLを表示」→ 回答用URLを業者さんに伝える
 *
 * 【手で直さないもの】
 *  「JWNET出力」シートの2行目の関数、「フォーム回答」の「出力状況」列
 */

const APP = {
  FORM_TITLE: 'マニフェスト登録入力（収集運搬業者様用）',
  SHEET_RESP: 'フォーム回答',
  SHEET_OUT: 'JWNET出力',
  SHEET_SET: '設定',
  SHEET_CTRL: '出力管理',
  PROP_FORM_ID: 'FORM_ID',
  DATA_START: 3,      // 設定シート：データ開始行（1行目=ブロック名、2行目=見出し）
  MAX_SET_ROWS: 300,  // 設定シート：読み取る最大行
  OUT_ROWS: 3000,     // JWNET出力シートの行数（一度に展開できる最大件数）
  CSV_FOLDER: 'JWNET_CSV出力',
  OTHER_SITE: '【リストにない現場】',
  LABEL_APPROVE: '確認（✔で承認）',
  LABEL_STATUS: '出力状況',
  ST_SAVED: 'CSV保存済',
  ST_DONE: '取込済',
};

// 野村工務店さん側の確認（✔）を必須にするか。外部の業者さんが直接入力するので、通常は true。
const REQUIRE_APPROVAL = true;

// JWNET「新規登録（読込/保存）」75列の見出し（template-v2.csv と同じ並び・同じ表記）
const JWNET_HEADERS = [
  '連絡番号1', '連絡番号2', '連絡番号3', '引渡し日', '排出事業場（コード）',
  '引渡し担当者', '登録担当者', '廃棄物の種類（コード）', '廃棄物の名称', '廃棄物の数量',
  '廃棄物の数量単位（コード）', '荷姿（コード）', '荷姿の数量', '数量の確定者（コード）', '有害物質1（コード）',
  '有害物質2（コード）', '有害物質3（コード）', '有害物質4（コード）', '有害物質5（コード）', '有害物質6（コード）',
  '[区間1] 収集運搬業者加入者番号', '[区間1] 再委託収集運搬業者加入者番号', '[区間1] 運搬方法（コード）', '[区間1] 車両番号', '[区間1] 運搬担当者',
  '[区間1] 運搬先事業場加入者番号', '[区間1] 運搬先事業場（コード）', '[区間2] 収集運搬業者加入者番号', '[区間2] 再委託収集運搬業者加入者番号', '[区間2] 運搬方法（コード）',
  '[区間2] 車両番号', '[区間2] 運搬担当者', '[区間2] 運搬先事業場加入者番号', '[区間2] 運搬先事業場（コード）', '[区間3] 収集運搬業者加入者番号',
  '[区間3] 再委託収集運搬業者加入者番号', '[区間3] 運搬方法（コード）', '[区間3] 車両番号', '[区間3] 運搬担当者', '[区間3] 運搬先事業場加入者番号',
  '[区間3] 運搬先事業場（コード）', '[区間4] 収集運搬業者加入者番号', '[区間4] 再委託収集運搬業者加入者番号', '[区間4] 運搬方法（コード）', '[区間4] 車両番号',
  '[区間4] 運搬担当者', '[区間4] 運搬先事業場加入者番号', '[区間4] 運搬先事業場（コード）', '[区間5] 収集運搬業者加入者番号', '[区間5] 再委託収集運搬業者加入者番号',
  '[区間5] 運搬方法（コード）', '[区間5] 車両番号', '[区間5] 運搬担当者', '[区間5] 運搬先事業場加入者番号', '[区間5] 運搬先事業場（コード）',
  '処分業者加入者番号', '再委託処分業者加入者番号', '処分事業場（コード）', '処分方法（コード）', '最終処分事業場登録区分',
  '最終処分事業場1（コード）', '最終処分事業場2（コード）', '最終処分事業場3（コード）', '最終処分事業場4（コード）', '最終処分事業場5（コード）',
  '最終処分事業場6（コード）', '最終処分事業場7（コード）', '最終処分事業場8（コード）', '最終処分事業場9（コード）', '最終処分事業場10（コード）',
  '備考1', '備考2', '備考3', '備考4', '備考5'
];

// 任意の質問をフォームに入れるかどうか。変えたら、新しいスプレッドシートで setupAll をやり直す
//  ※「必須項目だが、業者さんが答えられるか分からないもの」は、任意の質問として残してある。
//    答えがあればそれを使い、空欄なら設定シートの既定値で補う。誰が入力するか決まったら、
//    フォームの編集画面で「必須」にするか、ここを false にして作り直す。
const OPTIONAL = {
  renraku1: true,    // 貴社の伝票番号（→ 連絡番号1）
  handover: true,    // 引渡し担当者（空欄なら、現場マスタの担当者）【JWNETの必須項目】
  fixer: true,       // 数量の確定者（空欄なら、固定値）【JWNETの必須項目】
  finalkind: true,   // 最終処分事業場の登録区分（空欄なら、固定値）【JWNETの必須項目】
  final1: true,      // 最終処分事業場（登録区分が「当欄記載のとおり」のとき必須）
  registrant: false, // 登録担当者（空欄なら、固定値）
  wname: false,      // 廃棄物の名称
  pkgqty: false,     // 荷姿の数量
  vehicle: false,    // 車両番号（区間1）
  driver: false,     // 運搬担当者（区間1）
  remark1: false,    // 連絡事項（→ 備考1。JWNET上で他社にも見える）
};

// 設定シートのブロック（左から順に並ぶ）。1列目=フォームに出る名前、2列目以降=JWNETに渡すコードなど
// ※ seed の「（仮）」は必ず本物に書き換えること（JWNETの基本設定・各マスタで確認）
const BLOCKS = [
  { key: 'site',    label: '現場（排出事業場）', cols: ['現場名', '排出事業場コード（10桁以内）', '引渡し担当者の既定値（野村工務店の担当・全角12文字以内）'],
    seed: [['（仮）A現場', 'A010000001', '（仮）担当者']] },
  { key: 'waste',   label: '廃棄物の種類',       cols: ['廃棄物の種類名', '廃棄物の分類コード（7桁）'],              seed: [['（仮）廃棄物の種類', '0605000']] },
  { key: 'unit',    label: '数量の単位',         cols: ['単位名', '数量単位コード（1桁）'],                          seed: [['（仮）単位', '1']] },
  { key: 'pack',    label: '荷姿',               cols: ['荷姿名', '荷姿コード（2桁）'],                              seed: [['（仮）荷姿', '01']] },
  { key: 'fixer',   label: '数量の確定者',       cols: ['確定者名', '確定者コード（2桁）'],                          seed: [['（仮）確定者', '01']] },
  { key: 'carrier', label: '収集運搬業者',       cols: ['収集運搬業者名（フォームの「貴社名」）', '加入者番号（7桁）'], seed: [['（仮）運搬会社', '1234567']] },
  { key: 'dispo',   label: '処分先（業者＋事業場）', cols: ['処分先名（業者／事業場）', '処分業者 加入者番号（7桁）', '処分事業場コード（3桁）'],
    seed: [['（仮）処分会社 ○○工場', '7654321', '001']] },
  { key: 'finalkind', label: '最終処分事業場の登録区分', cols: ['登録区分名', '登録区分コード（0または1）'],
    seed: [['委託契約書記載のとおり', '0'], ['当欄記載のとおり', '1']] },
  { key: 'final1',  label: '最終処分事業場',     cols: ['最終処分事業場名', '最終処分事業場コード（10桁以内）'],      seed: [['（仮）最終処分場', 'F010000001']] },
  { key: 'defaults', label: '固定値（フォームの答えが空欄のときに使う既定値）', cols: ['項目（変えないでください）', '値'],
    seed: [['登録担当者', ''], ['数量の確定者コード', '01'], ['最終処分事業場の登録区分コード', '0']] },
];

const ASK_HELP = '分かる場合のみ入力してください。空欄の場合は、野村工務店側で補います。';

// フォームの質問 → JWNET出力の列（col = 75列中の何列目か）
//  how: text=そのまま / date=yyyy/mm/dd / qty=0.000形式 / lookup=設定シートから名前→コード
//       lookupOpt=lookupだが未回答なら空欄 / lookupOr=lookupだが未回答なら固定値 / textOr=そのままだが未回答なら既定値
//  optional: OPTIONAL で入れる／入れないを切り替える質問　req:false: 常にあるが必須ではない質問
//  unless: その質問が入っているときは、この列の式は質問側が持つので作らない
const QUESTIONS = [
  { key: 'carrier', title: '貴社名', kind: 'list', block: 'carrier', outs: [{ col: 21, how: 'lookup', idx: 2 }],
    help: '収集運搬業者としての貴社名を選んでください' },
  { key: 'submitter', title: '入力者のお名前', kind: 'text', val: { maxlen: 20, help: '全角20文字まで' }, outs: [],
    help: '入力内容の確認でご連絡することがあります' },
  { key: 'renraku1', optional: true, req: false, title: '貴社の伝票番号', kind: 'text', textCol: true,
    val: { pattern: '^[!-~]{1,20}$', help: '半角の英数字・記号で20文字まで' }, outs: [{ col: 1, how: 'text' }],
    help: '貴社の伝票番号・管理番号があれば入力してください（なければ空欄で構いません）' },
  { key: 'date', title: '収集（引渡し）日', kind: 'date', outs: [{ col: 4, how: 'date' }], help: '廃棄物を引き取った日' },
  { key: 'site', title: '現場名', kind: 'list', block: 'site', extraChoices: [APP.OTHER_SITE],
    outs: [{ col: 5, how: 'lookup', idx: 2 }, { col: 6, how: 'lookup', idx: 3, unless: 'handover' }],
    help: 'リストにない現場のときは「' + APP.OTHER_SITE + '」を選び、次の欄に現場名を入力してください' },
  { key: 'siteother', req: false, title: '現場名（リストにない場合のみ）', kind: 'text', val: { maxlen: 40, help: '全角40文字まで' }, outs: [],
    help: '上で「' + APP.OTHER_SITE + '」を選んだときだけ入力してください' },
  { key: 'handover', optional: true, req: false, title: '引渡し担当者（野村工務店の方）', kind: 'text', val: { maxlen: 12, help: '全角12文字まで' },
    outs: [{ col: 6, how: 'textOr', fallback: 'site' }], help: ASK_HELP },
  { key: 'registrant', optional: true, req: false, title: '登録担当者', kind: 'text', val: { maxlen: 12, help: '全角12文字まで' },
    outs: [{ col: 7, how: 'textOr', fallback: 'const', label: '登録担当者' }], help: ASK_HELP },
  { key: 'waste', title: '廃棄物の種類', kind: 'list', block: 'waste', outs: [{ col: 8, how: 'lookup', idx: 2 }],
    help: '種類が複数あるときは、種類ごとに1回ずつ送信してください' },
  { key: 'wname', optional: true, req: false, title: '廃棄物の名称', kind: 'text', textCol: true,
    val: { maxlen: 30, help: '全角30文字まで' }, outs: [{ col: 9, how: 'text' }] },
  { key: 'qty', title: '廃棄物の数量', kind: 'text', textCol: true,
    val: { pattern: '^[0-9]{1,5}(\\.[0-9]{1,3})?$', help: '半角数字。整数は5桁まで、小数は3桁まで（例：5.5）' },
    outs: [{ col: 10, how: 'qty' }] },
  { key: 'unit', title: '数量の単位', kind: 'list', block: 'unit', outs: [{ col: 11, how: 'lookup', idx: 2 }] },
  { key: 'pack', title: '荷姿', kind: 'list', block: 'pack', outs: [{ col: 12, how: 'lookup', idx: 2 }] },
  { key: 'pkgqty', optional: true, req: false, title: '荷姿の数量', kind: 'text', textCol: true,
    val: { pattern: '^[0-9]{1,5}$', help: '半角数字で5桁まで' }, outs: [{ col: 13, how: 'text' }] },
  { key: 'fixer', optional: true, req: false, title: '数量の確定者', kind: 'list', block: 'fixer',
    outs: [{ col: 14, how: 'lookupOr', idx: 2, label: '数量の確定者コード' }], help: ASK_HELP },
  { key: 'vehicle', optional: true, req: false, title: '車両番号', kind: 'text', textCol: true,
    val: { maxlen: 15, help: '全角15文字まで' }, outs: [{ col: 24, how: 'text' }] },
  { key: 'driver', optional: true, req: false, title: '運搬担当者', kind: 'text', textCol: true,
    val: { maxlen: 12, help: '全角12文字まで' }, outs: [{ col: 25, how: 'text' }] },
  { key: 'dispo', title: '処分先', kind: 'list', block: 'dispo', outs: [{ col: 56, how: 'lookup', idx: 2 }, { col: 58, how: 'lookup', idx: 3 }],
    help: '処分業者と処分事業場の組み合わせを選んでください' },
  { key: 'finalkind', optional: true, req: false, title: '最終処分事業場の登録区分', kind: 'list', block: 'finalkind',
    outs: [{ col: 60, how: 'lookupOr', idx: 2, label: '最終処分事業場の登録区分コード' }], help: ASK_HELP },
  { key: 'final1', optional: true, req: false, title: '最終処分事業場', kind: 'list', block: 'final1',
    outs: [{ col: 61, how: 'lookupOpt', idx: 2 }], help: '上で「当欄記載のとおり」を選んだときだけ入力してください' },
  { key: 'remark1', optional: true, req: false, title: '連絡事項', kind: 'text', textCol: true,
    val: { maxlen: 25, help: '全角25文字まで' }, outs: [{ col: 71, how: 'text' }] },
];

// フォームの答えが無いときに、設定シートの「固定値」から入れる列（質問が入っているときは、質問側の式が固定値も使う）
const CONST_OUTS = [
  { col: 7, label: '登録担当者', unless: 'registrant' },
  { col: 14, label: '数量の確定者コード', unless: 'fixer' },
  { col: 60, label: '最終処分事業場の登録区分コード', unless: 'finalkind' },
];

// 必須の列（新規登録レイアウトの◎。運搬先事業場は1区間なので対象外）
const REQUIRED_COLS = [4, 5, 6, 8, 10, 11, 12, 14, 21, 56, 58, 60];

// 桁数チェック（JWNET出力の列番号 → 上限byte）。全角=2byte、半角=1byte
const BYTE_LIMITS = {
  1: 20, 2: 20, 3: 20, 4: 10, 5: 10, 6: 24, 7: 24, 8: 7, 9: 60, 10: 9, 11: 1, 12: 2, 13: 5, 14: 2,
  21: 7, 24: 30, 25: 24, 56: 7, 58: 3, 60: 1, 61: 10, 71: 50,
};

// ------------------------------------------------------------
// 部品（計算だけ。シートやフォームには触らない）
// ------------------------------------------------------------

function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function isRequired_(q) { return !q.optional && q.req !== false; }

function layoutBlocks_() {
  const layout = {};
  let start = 1;
  BLOCKS.forEach(function (b) {
    layout[b.key] = { start: start, width: b.cols.length, block: b };
    start += b.cols.length + 1; // 1列あけて次のブロック
  });
  return layout;
}

// ※回答シートには、回答が入るときに行が挿入されて「2行目から」の参照がずれることがある。
//   そのため行番号は書かず、列全体（$A:$A）を参照する。見出し行は ISNUMBER で除外する。
function respRef_(letter) { return "'" + APP.SHEET_RESP + "'!$" + letter + ':$' + letter; }

// 出力対象 ＝ 回答がある行 ＋（確認の✔が付いている）＋ まだ「取込済」でない行
function filt_(schema, letter) {
  const args = ['ISNUMBER(' + respRef_('A') + ')'];
  if (schema.approvalCol) args.push(respRef_(colLetter_(schema.approvalCol)) + '=TRUE');
  args.push(respRef_(colLetter_(schema.statusCol)) + '<>"' + APP.ST_DONE + '"');
  return 'FILTER(' + respRef_(letter) + ',' + args.join(',') + ')';
}

function countTargets_(schema) {
  const parts = [respRef_('A') + ',">0"'];
  if (schema.approvalCol) parts.push(respRef_(colLetter_(schema.approvalCol)) + ',TRUE');
  parts.push(respRef_(colLetter_(schema.statusCol)) + ',"<>' + APP.ST_DONE + '"');
  return 'COUNTIFS(' + parts.join(',') + ')';
}

function guard_(schema, inner) {
  return '=IF(' + countTargets_(schema) + '=0,"",' + inner + ')';
}

function setRange_(layout, key, lastIdx) {
  const b = layout[key];
  return "'" + APP.SHEET_SET + "'!$" + colLetter_(b.start) + ':$' + colLetter_(b.start + lastIdx - 1); // 列全体（行の挿入・削除に強い）
}

// 設定シートから「名前 → idx列目の値」を引く。引いた値が空欄なら空欄、名前が見つからなければ※未登録
function lookupExpr_(schema, keyExpr, block, idx) {
  const v = 'VLOOKUP(' + keyExpr + ',' + setRange_(schema.layout, block, idx) + ',' + idx + ',FALSE)';
  return 'IFERROR(IF(' + v + '="","",' + v + '),"※未登録")';
}

// 設定シートの「固定値」から、項目名で値を引く
function constExpr_(schema, label) {
  const v = 'VLOOKUP("' + label + '",' + setRange_(schema.layout, 'defaults', 2) + ',2,FALSE)';
  return 'IFERROR(IF(' + v + '="","",' + v + '),"※未登録")';
}

function outputFormula_(schema, q, out) {
  const f = filt_(schema, colLetter_(schema.respCols[q.key]));
  switch (out.how) {
    case 'text':
      return guard_(schema, 'ARRAYFORMULA(' + f + '&"")'); // 空欄は空欄のまま（0や#N/Aにしない）
    case 'date':
      return guard_(schema, 'ARRAYFORMULA(TEXT(' + f + ',"yyyy/mm/dd"))');
    case 'qty':
      return guard_(schema, 'ARRAYFORMULA(IFERROR(TEXT(VALUE(' + f + '),"0.000"),"※数量不正"))');
    case 'lookup':
      return guard_(schema, 'ARRAYFORMULA(' + lookupExpr_(schema, f, q.block, out.idx) + ')');
    case 'lookupOpt': // 未回答なら空欄
      return guard_(schema, 'ARRAYFORMULA(IF(' + f + '&""="","",' + lookupExpr_(schema, f, q.block, out.idx) + '))');
    case 'lookupOr': // 未回答なら、設定シートの固定値
      return guard_(schema, 'ARRAYFORMULA(IF(' + f + '&""<>"",' + lookupExpr_(schema, f, q.block, out.idx) + ',' + constExpr_(schema, out.label) + '))');
    case 'textOr': { // 未回答なら、現場マスタの既定値（fallback:site）か、固定値（fallback:const）
      const fb = out.fallback === 'site'
        ? lookupExpr_(schema, filt_(schema, colLetter_(schema.respCols.site)), 'site', 3)
        : constExpr_(schema, out.label);
      return guard_(schema, 'ARRAYFORMULA(IF(' + f + '&""<>"",' + f + '&"",' + fb + '))');
    }
    default:
      throw new Error('未対応のhow: ' + out.how);
  }
}

function constFormula_(schema, c) {
  const dateF = filt_(schema, colLetter_(schema.respCols.date));
  return guard_(schema, 'ARRAYFORMULA(IF(' + dateF + '<>"",' + constExpr_(schema, c.label) + ',""))');
}

function buildSchema_() {
  const questions = QUESTIONS.filter(function (q) { return !q.optional || OPTIONAL[q.key]; });
  const layout = layoutBlocks_();
  const respCols = {};
  questions.forEach(function (q, i) { respCols[q.key] = 2 + i; }); // A列はタイムスタンプ
  const approvalCol = REQUIRE_APPROVAL ? 2 + questions.length : null;
  const statusCol = 2 + questions.length + (REQUIRE_APPROVAL ? 1 : 0);
  const schema = { questions: questions, layout: layout, respCols: respCols, approvalCol: approvalCol, statusCol: statusCol, formulas: {} };
  function put(col, f) {
    if (schema.formulas[col]) throw new Error('JWNET出力の列が重複しています: ' + col);
    schema.formulas[col] = f;
  }
  const asked = {};
  questions.forEach(function (q) { asked[q.key] = true; });
  const skip = function (o) { return o.unless && asked[o.unless]; };
  questions.forEach(function (q) { q.outs.forEach(function (o) { if (!skip(o)) put(o.col, outputFormula_(schema, q, o)); }); });
  CONST_OUTS.forEach(function (c) { if (!skip(c)) put(c.col, constFormula_(schema, c)); });
  return schema;
}

function responseTitles_(schema) {
  const t = ['タイムスタンプ'].concat(schema.questions.map(function (q) { return q.title; }));
  if (schema.approvalCol) t.push(APP.LABEL_APPROVE);
  t.push(APP.LABEL_STATUS);
  return t;
}

function controlSheetFormulas_(schema) {
  const r = "'" + APP.SHEET_RESP + "'";
  const o = "'" + APP.SHEET_OUT + "'";
  const lastCol = colLetter_(JWNET_HEADERS.length);
  const conds = responseTitles_(schema).map(function (t, i) { return r + '!' + colLetter_(i + 1) + '1="' + t + '"'; });
  const bytes = Object.keys(BYTE_LIMITS).map(function (c) {
    const L = colLetter_(Number(c));
    return 'SUMPRODUCT(--(LENB(' + o + '!$' + L + '$2:$' + L + ')>' + BYTE_LIMITS[c] + '))';
  }).join('+');
  const blanks = REQUIRED_COLS.filter(function (c) { return c !== 5; }).map(function (c) {
    const L = colLetter_(c);
    return 'SUMPRODUCT((LEN(' + o + '!$E$2:$E)>0)*(LEN(' + o + '!$' + L + '$2:$' + L + ')=0))';
  }).join('+');
  // 確認待ち ＝（回答あり・未取込）−（回答あり・✔あり・未取込）。「<>TRUE」という条件の解釈に頼らない
  const notDone = respRef_(colLetter_(schema.statusCol)) + ',"<>' + APP.ST_DONE + '"';
  // 登録区分が「1（当欄記載のとおり）」の行は、最終処分事業場（61列目）が必須
  const kindL = colLetter_(60), finalL = colLetter_(61);
  const finalNeeded = 'SUMPRODUCT((' + o + '!$' + kindL + '$2:$' + kindL + '="1")*(LEN(' + o + '!$' + finalL + '$2:$' + finalL + ')=0))';
  const waiting = schema.approvalCol
    ? '=COUNTIFS(' + respRef_('A') + ',">0",' + notDone + ')-' + countTargets_(schema)
    : '="（確認機能はオフです）"';
  return {
    B3: waiting,
    B4: '=' + countTargets_(schema),
    B5: '=SUMPRODUCT(--(LEN(' + o + '!$E$2:$E)>0))',
    B6: '=IF($B$4=$B$5,"OK","NG：確認済みの回答の件数と、出力の件数が違います")',
    B7: '=IF(AND(' + conds.join(',') + '),"OK","NG：フォームの質問や回答シートの列が追加・変更・並べ替えされています")',
    B8: '=IF(COUNTIF(' + o + '!$A$2:$' + lastCol + ',"※*")=0,"OK","NG：設定シートにない名前が選ばれています（※が出ているセルを確認）")',
    B9: '=IFERROR(IF((' + bytes + ')+(' + blanks + ')+(' + finalNeeded + ')=0,"OK","NG：桁数（byte）オーバーか、必須項目の空欄があります（登録区分が「当欄記載のとおり」の行は、最終処分事業場も必須）"),"NG：JWNET出力シートにエラーがあります")',
    B10: '=IF(AND($B$4>0,$B$6="OK",$B$7="OK",$B$8="OK",$B$9="OK"),"OK","NG")',
  };
}

function csvEscape_(v) {
  const s = String(v == null ? '' : v);
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function toCsv_(rows) {
  return rows.map(function (r) { return r.map(csvEscape_).join(','); }).join('\r\n') + '\r\n';
}

function readBlockNames_(setSheet, entry) {
  const vals = setSheet.getRange(APP.DATA_START, entry.start, APP.MAX_SET_ROWS - APP.DATA_START + 1, 1).getValues();
  return vals.map(function (r) { return String(r[0]).trim(); }).filter(function (s) { return s !== ''; });
}

function findDuplicates_(arr) {
  const seen = {}, dup = [];
  arr.forEach(function (s) { if (seen[s] && dup.indexOf(s) < 0) dup.push(s); seen[s] = true; });
  return dup;
}

// 回答シートから、出力対象の行（回答あり＋確認済み＋取込済でない）を、シートの並び順で返す
function readTargets_(resp, schema) {
  const last = resp.getLastRow();
  if (last < 2) return [];
  const vals = resp.getRange(2, 1, last - 1, schema.statusCol).getValues();
  const out = [];
  vals.forEach(function (v, i) {
    const approved = schema.approvalCol ? v[schema.approvalCol - 1] === true : true;
    if (v[0] instanceof Date && approved && String(v[schema.statusCol - 1]) !== APP.ST_DONE) out.push({ row: i + 2, ts: v[0].getTime() });
  });
  return out;
}

// 出力状況を書く。行番号ではなく、今あらためて読み直した回答の日時で突き合わせる（行がずれても間違えない）
function markStatus_(resp, schema, stamps, fromStatus, toStatus) {
  const set = {};
  stamps.forEach(function (t) { set[t] = true; });
  const last = resp.getLastRow();
  if (last < 2) return 0;
  const vals = resp.getRange(2, 1, last - 1, schema.statusCol).getValues();
  let n = 0;
  vals.forEach(function (v, i) {
    const approved = schema.approvalCol ? v[schema.approvalCol - 1] === true : true;
    const st = String(v[schema.statusCol - 1]);
    if (v[0] instanceof Date && approved && set[v[0].getTime()] && st !== APP.ST_DONE && (fromStatus === null || st === fromStatus)) {
      resp.getRange(i + 2, schema.statusCol).setValue(toStatus);
      n++;
    }
  });
  return n;
}

// ------------------------------------------------------------
// セットアップ（初回に1回だけ実行）
// ------------------------------------------------------------

function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートの「拡張機能 > Apps Script」から開いたスクリプトで実行してください。');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(APP.PROP_FORM_ID)) {
    throw new Error('すでにセットアップ済みです。やり直す場合は、新しいスプレッドシートで最初から実行してください。');
  }
  [APP.SHEET_RESP, APP.SHEET_OUT, APP.SHEET_SET, APP.SHEET_CTRL].forEach(function (n) {
    if (ss.getSheetByName(n)) throw new Error('「' + n + '」シートがすでにあります。新しいスプレッドシートで実行してください。');
  });

  ss.setSpreadsheetLocale('ja_JP');
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  const schema = buildSchema_();

  // 1) 設定シート
  const setSheet = ss.insertSheet(APP.SHEET_SET);
  writeSettings_(setSheet, schema);

  // 2) フォーム
  const form = createForm_(schema, setSheet);
  props.setProperty(APP.PROP_FORM_ID, form.getId());

  // 3) 回答シート（1枚目）
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  const resp = findResponseSheet_(ss);
  resp.setName(APP.SHEET_RESP);
  ss.setActiveSheet(resp);
  ss.moveActiveSheet(1);
  waitForHeaders_(resp, schema);
  prepareResponseSheet_(resp, schema);

  // 4) JWNET出力シート（2枚目）
  const outSheet = ss.insertSheet(APP.SHEET_OUT, 1);
  writeOutputSheet_(outSheet, schema);

  // 5) 出力管理シート（4枚目）
  const ctrl = ss.insertSheet(APP.SHEET_CTRL);
  writeControlSheet_(ctrl, schema, form);

  // 6) 最初からあった空のシートを消す
  ['シート1', 'Sheet1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  ss.setActiveSheet(ss.getSheetByName(APP.SHEET_SET));
  Logger.log('★ここまで成功しました。');
  Logger.log('入力フォーム（業者さんに伝えるURL）: ' + form.getPublishedUrl());
  Logger.log('入力フォーム（編集用・管理者だけ）: ' + form.getEditUrl());
  Logger.log('次は「設定」シートの（仮）を本物に書き換え、メニュー JWNET > フォームの選択肢を更新 を押してください。');
}

function writeSettings_(sh, schema) {
  const layout = schema.layout;
  const last = BLOCKS.reduce(function (n, b) { return layout[b.key].start + b.cols.length - 1; }, 0);
  if (sh.getMaxColumns() < last) sh.insertColumnsAfter(sh.getMaxColumns(), last - sh.getMaxColumns());
  if (sh.getMaxRows() < APP.MAX_SET_ROWS) sh.insertRowsAfter(sh.getMaxRows(), APP.MAX_SET_ROWS - sh.getMaxRows());
  // コードの先頭の0（例：001）が消えないよう、全部「書式なしテキスト」にしておく
  sh.getRange(1, 1, APP.MAX_SET_ROWS, last).setNumberFormat('@');
  BLOCKS.forEach(function (b) {
    const e = layout[b.key];
    sh.getRange(1, e.start).setValue('▼ ' + b.label).setFontWeight('bold');
    sh.getRange(2, e.start, 1, e.width).setValues([b.cols]).setFontWeight('bold').setBackground('#e8f0fe').setWrap(true);
    if (b.seed.length) sh.getRange(APP.DATA_START, e.start, b.seed.length, e.width).setValues(b.seed);
    sh.setColumnWidth(e.start, 220);
  });
  sh.setFrozenRows(2);
}

function createForm_(schema, setSheet) {
  const form = FormApp.create(APP.FORM_TITLE);
  form.setDescription(
    '収集運搬業者様用の入力フォームです。FAXの代わりに、こちらから入力してください。\n' +
    '・1回の送信で、1つの廃棄物（種類）を入力します。種類が複数あるときは、種類ごとに送信してください。\n' +
    '・「分かる場合のみ」と書かれた項目は、分からなければ空欄で構いません（野村工務店側で補います）。\n' +
    '・入力内容は、野村工務店で確認のうえ、電子マニフェスト（JWNET）に登録されます。');
  form.setCollectEmail(false); // 回答シートの列がずれるため使わない
  form.setAllowResponseEdits(false);
  form.setShowLinkToRespondAgain(true);
  form.setConfirmationMessage('送信しました。続けて入力する場合は「別の回答を送信」を押してください。');
  try { form.setRequireLogin(false); } catch (e) { /* 個人アカウントでは設定項目がない場合がある */ }
  schema.questions.forEach(function (q) {
    const req = isRequired_(q);
    if (q.kind === 'date') {
      const it = form.addDateItem().setTitle(q.title).setRequired(req).setIncludesYear(true);
      if (q.help) it.setHelpText(q.help);
    } else if (q.kind === 'list') {
      const names = readBlockNames_(setSheet, schema.layout[q.block]).concat(q.extraChoices || []);
      if (!names.length) throw new Error('設定シートに「' + q.title + '」の選択肢がありません。');
      const it = form.addListItem().setTitle(q.title).setRequired(req).setChoiceValues(names);
      if (q.help) it.setHelpText(q.help);
    } else if (q.kind === 'text') {
      const it = form.addTextItem().setTitle(q.title).setRequired(req);
      it.setValidation(buildValidation_(q.val));
      it.setHelpText(q.help ? q.help + '（' + q.val.help + '）' : q.val.help);
    } else {
      throw new Error('未対応のkind: ' + q.kind);
    }
  });
  return form;
}

function buildValidation_(v) {
  let b = FormApp.createTextValidation();
  b = v.pattern ? b.requireTextMatchesPattern(v.pattern) : b.requireTextLengthLessThanOrEqualTo(v.maxlen);
  return b.setHelpText(v.help).build();
}

function findResponseSheet_(ss) {
  for (let i = 0; i < 10; i++) {
    SpreadsheetApp.flush();
    const fresh = SpreadsheetApp.openById(ss.getId());
    const found = fresh.getSheets().filter(function (s) { return s.getFormUrl(); })[0];
    if (found) return found;
    Utilities.sleep(1000);
  }
  throw new Error('フォームの回答シートが見つかりません。新しいスプレッドシートで最初からやり直してください。');
}

function waitForHeaders_(resp, schema) {
  const expected = ['タイムスタンプ'].concat(schema.questions.map(function (q) { return q.title; }));
  for (let i = 0; i < 10; i++) {
    const got = resp.getRange(1, 1, 1, expected.length).getValues()[0];
    if (expected.every(function (t, k) { return got[k] === t; })) return;
    Utilities.sleep(1000);
    SpreadsheetApp.flush();
  }
  throw new Error('回答シートの見出しがフォームの質問と一致しません。新しいスプレッドシートで最初からやり直してください。');
}

// 回答シートの右端に「確認」列と「出力状況」列を足し、文字列にしたい列の書式を整える
function prepareResponseSheet_(resp, schema) {
  if (resp.getMaxColumns() < schema.statusCol) resp.insertColumnsAfter(resp.getMaxColumns(), schema.statusCol - resp.getMaxColumns());
  const rows = resp.getMaxRows() - 1;
  if (schema.approvalCol) {
    resp.getRange(1, schema.approvalCol).setValue(APP.LABEL_APPROVE).setFontWeight('bold').setBackground('#fff2cc');
    resp.getRange(2, schema.approvalCol, rows, 1).insertCheckboxes();
  }
  resp.getRange(1, schema.statusCol).setValue(APP.LABEL_STATUS).setFontWeight('bold').setBackground('#eeeeee');
  resp.getRange(2, schema.statusCol, rows, 1).setNumberFormat('@');
  schema.questions.forEach(function (q) {
    if (q.textCol) { // 先頭の0が消えないよう、文字列の列にしておく
      resp.getRange(2, schema.respCols[q.key], rows, 1).setNumberFormat('@');
    }
  });
}

// すでにセットアップ済みのシートで、関数だけを最新版に入れ直す（データ・設定・フォームは触らない）
function rewriteFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!PropertiesService.getScriptProperties().getProperty(APP.PROP_FORM_ID)) {
    throw new Error('先に setupAll を実行してください。');
  }
  const schema = buildSchema_();
  const outSheet = ss.getSheetByName(APP.SHEET_OUT);
  const ctrl = ss.getSheetByName(APP.SHEET_CTRL);
  if (!outSheet || !ctrl) throw new Error('「' + APP.SHEET_OUT + '」または「' + APP.SHEET_CTRL + '」シートが見つかりません。');
  Object.keys(schema.formulas).forEach(function (c) {
    outSheet.getRange(2, Number(c)).setFormula(schema.formulas[c]);
  });
  const f = controlSheetFormulas_(schema);
  Object.keys(f).forEach(function (a1) { ctrl.getRange(a1).setFormula(f[a1]); });
  SpreadsheetApp.flush();
  Logger.log('★関数を入れ直しました（JWNET出力 ' + Object.keys(schema.formulas).length + '列、出力管理 ' + Object.keys(f).length + 'か所）。');
}

function writeOutputSheet_(sh, schema) {
  const n = JWNET_HEADERS.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  if (sh.getMaxRows() < APP.OUT_ROWS) sh.insertRowsAfter(sh.getMaxRows(), APP.OUT_ROWS - sh.getMaxRows());
  sh.getRange(1, 1, 1, n).setValues([JWNET_HEADERS]).setFontWeight('bold').setBackground('#e8f0fe');
  Object.keys(schema.formulas).forEach(function (c) {
    sh.getRange(2, Number(c)).setFormula(schema.formulas[c]);
  });
  sh.setFrozenRows(1);
  sh.protect().setWarningOnly(true).setDescription('自動展開のシートです。直接編集しないでください。');
}

function writeControlSheet_(sh, schema, form) {
  const f = controlSheetFormulas_(schema);
  sh.getRange('A1').setValue('出力管理（B10 が OK のときだけ、CSVを出力してJWNETに取り込む）').setFontWeight('bold');
  sh.getRange('A2').setValue('入力フォーム（業者さんに伝えるURL）');
  sh.getRange('B2').setValue(form.getPublishedUrl());
  sh.getRange('A3').setValue('確認待ちの件数（まだ ✔ が付いていない回答）');
  sh.getRange('A4').setValue('出力対象の件数（✔済み・未取込）');
  sh.getRange('A5').setValue('JWNET出力シートの件数');
  sh.getRange('A6').setValue('件数の一致');
  sh.getRange('A7').setValue('フォームの質問と列の並び');
  sh.getRange('A8').setValue('設定にない名前の選択');
  sh.getRange('A9').setValue('桁数（byte）・必須の空欄');
  sh.getRange('A10').setValue('総合').setFontWeight('bold');
  Object.keys(f).forEach(function (a1) { sh.getRange(a1).setFormula(f[a1]); });
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('NG').setBackground('#f4c7c3').setRanges([sh.getRange('B6:B10')]).build();
  sh.setConditionalFormatRules([rule]);
  sh.getRange('A12').setValue('【使い方】').setFontWeight('bold');
  sh.getRange('A13').setValue('1. 業者さんがフォームに入力 →「フォーム回答」に溜まる（B3 に確認待ちの件数が出る）');
  sh.getRange('A14').setValue('2. 野村工務店で内容を確認し、問題なければ「' + APP.LABEL_APPROVE + '」列に ✔ を付ける');
  sh.getRange('A15').setValue('   ※現場名が「' + APP.OTHER_SITE + '」の行は、JWNETで排出事業場を登録し、設定シートに追加してから、現場名のセルを正しい名前に直して ✔');
  sh.getRange('A16').setValue('3. B10 が OK であること（NGなら B6〜B9 の理由を直す）');
  sh.getRange('A17').setValue('4. メニュー「JWNET > CSVを保存」→ JWNETの新規登録（読込）で取り込む');
  sh.getRange('A18').setValue('5. 取り込めたら、メニュー「JWNETに取り込めた（出力済みにする）」を押す（二重登録の防止）');
  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 380);
}

// ------------------------------------------------------------
// メニュー
// ------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi().createMenu('JWNET')
    .addItem('フォームの選択肢を更新', 'syncChoices')
    .addSeparator()
    .addItem('CSVを保存（Shift_JIS）', 'exportCsvSjis')
    .addItem('CSVを保存（UTF-8 BOM付き）', 'exportCsvUtf8')
    .addItem('JWNETに取り込めた（出力済みにする）', 'markExported')
    .addSeparator()
    .addItem('入力フォームのURLを表示', 'showFormUrl')
    .addToUi();
}

function showFormUrl() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.PROP_FORM_ID);
  const form = FormApp.openById(id);
  SpreadsheetApp.getUi().alert('入力フォーム',
    '回答用（業者さんに伝えるURL）:\n' + form.getPublishedUrl() + '\n\n編集用（管理者だけ。業者さんには伝えない）:\n' + form.getEditUrl(),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncChoices() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = PropertiesService.getScriptProperties().getProperty(APP.PROP_FORM_ID);
  if (!id) { ui.alert('先に setupAll を実行してください。'); return; }
  const form = FormApp.openById(id);
  const setSheet = ss.getSheetByName(APP.SHEET_SET);
  const schema = buildSchema_();
  const items = form.getItems(FormApp.ItemType.LIST);
  const problems = [];
  let updated = 0;
  schema.questions.filter(function (q) { return q.kind === 'list'; }).forEach(function (q) {
    const names = readBlockNames_(setSheet, schema.layout[q.block]);
    const item = items.filter(function (i) { return i.getTitle() === q.title; })[0];
    if (!item) { problems.push('フォームに「' + q.title + '」が見つかりません'); return; }
    if (!names.length) { problems.push('「' + q.title + '」の選択肢が空です（設定シート）'); return; }
    const dup = findDuplicates_(names.concat(q.extraChoices || []));
    if (dup.length) { problems.push('「' + q.title + '」に同じ名前があります：' + dup.join('、')); return; }
    item.asListItem().setChoiceValues(names.concat(q.extraChoices || []));
    updated++;
  });
  if (problems.length) {
    ui.alert('一部を更新できませんでした', problems.join('\n') + '\n\n（更新できた質問：' + updated + '件）', ui.ButtonSet.OK);
  } else {
    ui.alert('更新しました', '選択肢を更新しました（' + updated + '件）。', ui.ButtonSet.OK);
  }
}

// ------------------------------------------------------------
// CSV保存
// ------------------------------------------------------------

function exportCsvSjis() { exportCsv_('Shift_JIS'); }
function exportCsvUtf8() { exportCsv_('UTF-8'); }

function exportCsv_(charset) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctrl = ss.getSheetByName(APP.SHEET_CTRL);
  const resp = ss.getSheetByName(APP.SHEET_RESP);
  const outSheet = ss.getSheetByName(APP.SHEET_OUT);
  const schema = buildSchema_();

  // 1) チェック
  const st = ctrl.getRange('B4:B10').getDisplayValues().map(function (r) { return r[0]; }); // [対象件数, 出力件数, 一致, 列, 未登録, 桁数, 総合]
  if (st[6] !== 'OK') {
    const msgs = [st[2], st[3], st[4], st[5]].filter(function (s) { return s !== 'OK'; });
    if (Number(st[0]) === 0) msgs.unshift('出力する回答がありません（確認の ✔ が付いた回答が0件です）');
    ui.alert('CSVを保存できません', '「出力管理」シートを確認してください。\n\n' + msgs.join('\n'), ui.ButtonSet.OK);
    return;
  }

  // 2) 出力対象の回答を先に読む
  const targets = readTargets_(resp, schema);

  // 3) 出力シートを読む（見出し＋対象件数ぶん。その次の行が空であることも確かめる）
  const need = targets.length + 2;
  const readRows = Math.min(need, outSheet.getMaxRows());
  const all = outSheet.getRange(1, 1, readRows, JWNET_HEADERS.length).getDisplayValues();
  const rows = all.slice(0, targets.length + 1);
  const extra = all.length > targets.length + 1 ? all[targets.length + 1] : null;
  if (rows[0].join('\t') !== JWNET_HEADERS.join('\t')) { ui.alert('JWNET出力シートの見出しが変わっています。'); return; }
  const blankRow = function (r) { return r.every(function (v) { return v === ''; }); };
  const hasBlank = rows.slice(1).some(blankRow);
  if (targets.length === 0 || rows.length - 1 !== targets.length || hasBlank || (extra && !blankRow(extra))) {
    ui.alert('CSVを保存できません',
      '出力対象の回答は ' + targets.length + ' 件ですが、展開された行がそれと合いません。\n' +
      '入力や確認の途中だった可能性があります。少し待って、もう一度実行してください。', ui.ButtonSet.OK);
    return;
  }

  // 4) 保存（文字化けしないか読み戻して確認）
  const text = (charset === 'UTF-8' ? '\uFEFF' : '') + toCsv_(rows);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const name = 'jwnet_new_registration_' + stamp + (charset === 'UTF-8' ? '_utf8' : '_sjis') + '.csv';
  const blob = Utilities.newBlob('', 'text/csv', name);
  blob.setDataFromString(text, charset);
  if (blob.getDataAsString(charset) !== text) {
    ui.alert('CSVを保存できません',
      'この文字コードで表せない文字が含まれています（①や環境依存文字など）。\n設定シートや回答の文字を、ふつうの漢字・かなに直してください。', ui.ButtonSet.OK);
    return;
  }
  const it = DriveApp.getFoldersByName(APP.CSV_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(APP.CSV_FOLDER);
  const file = folder.createFile(blob);

  // 5) 回答シートの「出力状況」に「CSV保存済」と書く（JWNETに取り込めたあとで「取込済」にする）
  markStatus_(resp, schema, targets.map(function (t) { return t.ts; }), null, APP.ST_SAVED);
  ui.alert('保存しました',
    name + '（' + targets.length + '件）\n\n' + file.getUrl() + '\nフォルダ：マイドライブ / ' + APP.CSV_FOLDER +
    '\n\n【次にやること】\nJWNETで取り込み → 成功したら、メニュー「JWNET > JWNETに取り込めた（出力済みにする）」を押してください。\n' +
    '押すまでは、同じ回答が何度でも出力できます。', ui.ButtonSet.OK);
}

function markExported() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ss.getSheetByName(APP.SHEET_RESP);
  const schema = buildSchema_();
  const last = resp.getLastRow();
  const saved = [];
  if (last >= 2) {
    resp.getRange(2, 1, last - 1, schema.statusCol).getValues().forEach(function (v) {
      if (v[0] instanceof Date && String(v[schema.statusCol - 1]) === APP.ST_SAVED) saved.push(v[0].getTime());
    });
  }
  if (!saved.length) { ui.alert('出力済みにできるCSVがありません。先に「CSVを保存」をしてください。'); return; }
  const yn = ui.alert('出力済みにしますか？',
    '「' + APP.ST_SAVED + '」の ' + saved.length + ' 件を、JWNETに取り込めた場合だけ「はい」を押してください。\n' +
    '「はい」にすると、この ' + saved.length + ' 件は次回から出力されません。', ui.ButtonSet.YES_NO);
  if (yn !== ui.Button.YES) return;
  const n = markStatus_(resp, schema, saved, APP.ST_SAVED, APP.ST_DONE);
  ui.alert('出力済みにしました（' + n + '件）。');
}
