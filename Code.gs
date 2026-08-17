/**
 * 模擬試験アプリ GAS Web App（マルチユーザー対応版 / GAS配信方式）
 *
 * 画面(HTML)をGAS自身が配信し、データのやり取りは google.script.run で行う。
 * これにより file:// / CORS / Googleアカウント多重ログインの問題がすべて解消される。
 *
 * 使い方（全部入り1ファイル版）:
 *  1. コード.gsの中身を全部消し、このコードを貼り付けて保存
 *  2. [デプロイ]→[デプロイを管理]→鉛筆→バージョン「新バージョン」→[デプロイ]
 *  3. 表示された /exec のURLを開けば動作する（別のHTMLファイルは不要）
 *
 * ---- この版について ----
 * 2026/08/17 通常版（本番モード 50問・60分）
 * 変更1: 選択肢の取得を getChoicesBatch で1回にまとめ、出題準備を高速化
 * 変更2: 用語学習メニューに受験者名の表示と「別の学生に変更」を追加
 * 変更3: 用語学習から受験者を変更した場合、選択後は用語学習メニューへ戻るようにした
 * 変更4: 用語クイズの出題画面に用語ID（TERM-xxx）を表示（模擬試験の出題画面と同じ形）
 * 変更7: インポートのエラー表示を整理。IDを並べる代わりに
 *        「CSVの内容がそろっていません。ダウンロードし直して…」と、やることを1行で伝える形にした
 * 変更6: CSVの「対応試験」列を読み取り、プレビュー・取り込み完了の画面に表示。
 *        data_versionシートにも 'kintone' 固定ではなくCSVの値を記録するようにした
 * 変更5: 問題インポート画面を「CSVファイルを選ぶ」だけに整理（貼り付け欄と形式説明は廃止。形式は保守マニュアル側に記載）。
 *        読み込んだファイル名を画面に表示するようにした
 * 注: CSVのドラッグ＆ドロップはGASのサンドボックスiframe内では動作しないため見送り
 *     （インポートは「ファイル選択」ボタン、またはCSVの貼り付けで行う）
 */

// ===== グローバル設定 =====
var SS = SpreadsheetApp.getActiveSpreadsheet();

// シート名
var SHEET_QUESTIONS = 'questions';
var SHEET_CHOICES = 'choices';
var SHEET_HISTORY = 'history';
var SHEET_HISTORY_RESULTS = 'history_results';
var SHEET_TERMS = 'terms';
var SHEET_TERM_PROGRESS = 'term_progress';
var SHEET_TERM_QUIZ_HISTORY = 'term_quiz_history';
var SHEET_STUDENTS = 'students';

// ===== 必要なシート(タブ)を自動で用意する =====
function ensureSheets_() {
  var defs = {
    'questions': ['question_id', 'question_category', 'question_type', 'select_type', 'question'],
    'choices': ['question_id', 'choice_no', 'choice_text', 'explanation_text', 'explanation_url', 'is_correct'],
    'students': ['student_name', 'display_order'],
    'history': ['history_id', 'exam_datetime', 'student_name', 'exam_mode', 'selected_category', 'selected_type', 'total_count', 'total_score', 'correct_rate', 'duration_seconds'],
    'history_results': ['history_id', 'row_no', 'question_id', 'question_text', 'selected_answer', 'judgment', 'explanation_text', 'explanation_url'],
    'terms': ['term_id', 'chapter', 'term', 'definition', 'level', 'explanation_url'],
    'term_progress': ['term_id', 'student_name', 'learned_at'],
    'term_quiz_history': ['quiz_id', 'exam_datetime', 'student_name', 'term_id', 'chapter', 'level', 'judgment'],
    'data_version': ['version', 'target_exam', 'imported_at', 'imported_count']
  };
  Object.keys(defs).forEach(function (name) {
    var sh = SS.getSheetByName(name);
    if (!sh) {
      sh = SS.insertSheet(name);
      sh.getRange(1, 1, 1, defs[name].length).setValues([defs[name]]).setFontWeight('bold');
      if (name === 'students') { sh.appendRow(['濤川憲紀', 1]); } // 初期受験者（studentsタブで自由に追加・変更可）
    }
  });

  // term_quiz_history は「セッションごと集計」から「回答ごと1レコード」に方式変更した。
  // データが1件も入っていない場合に限り、見出し行を新方式へ書き換える（既存データは壊さない）。
  try {
    var tq = SS.getSheetByName('term_quiz_history');
    if (tq && tq.getLastRow() <= 1) {
      var want = defs['term_quiz_history'];
      var now = tq.getRange(1, 1, 1, Math.max(tq.getLastColumn(), want.length)).getValues()[0];
      var same = true;
      for (var wi = 0; wi < want.length; wi++) {
        if (String(now[wi] || '') !== want[wi]) { same = false; break; }
      }
      if (!same) {
        tq.clear();
        tq.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold');
      }
    }
  } catch (e2) {}
}

// ===== 画面配信 =====
function doGet(e) {
  try { ensureSheets_(); } catch (err) {}
  return HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>模擬試験アプリ</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html { scroll-behavior: smooth; }
        body { font-family: 'Segoe UI', 'Noto Sans JP', sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; display: flex; justify-content: center; align-items: center; padding: 20px; }
        .container { background: white; border-radius: 12px; box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3); max-width: 900px; width: 100%; padding: 40px; }
        h1 { color: #333; text-align: center; margin-bottom: 10px; font-size: 28px; }
        .subtitle { text-align: center; color: #666; font-size: 14px; margin-bottom: 30px; }
        .screen { display: none; }
        .screen.active { display: block; animation: fadeIn 0.3s ease-in; }
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        .name-select-screen { text-align: center; }
        .name-select-screen h2 { color: #667eea; margin-bottom: 20px; font-size: 20px; }
        .student-buttons { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .student-btn { padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }
        .student-btn:hover { transform: translateY(-3px); box-shadow: 0 10px 25px rgba(102, 126, 234, 0.4); }
        .student-info { background: #f8f9ff; padding: 15px; border-radius: 8px; margin-bottom: 25px; border-left: 4px solid #667eea; }
        .student-info p { color: #333; margin: 5px 0; }
        .student-info .student-name { font-size: 18px; font-weight: bold; color: #667eea; }
        .menu-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 20px; }
        .menu-item { background: #f9f9f9; border: 2px solid #e0e0e0; border-radius: 8px; padding: 20px; cursor: pointer; transition: all 0.3s; text-align: center; }
        .menu-item:hover { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-color: #667eea; color: white; transform: translateY(-5px); }
        .menu-item h3 { font-size: 18px; margin-bottom: 10px; }
        .menu-item p { font-size: 12px; opacity: 0.8; }
        /* ===== 問題インポート ===== */
        .import-help { margin-bottom: 14px; }
        .import-help details { background:#f8f9ff; border:1px solid #e0e0e0; border-radius:8px; padding:10px 14px; margin-bottom:10px; }
        .import-help summary { cursor:pointer; color:#667eea; font-weight:bold; font-size:14px; }
        .import-help-body { margin-top:10px; font-size:13px; color:#333; line-height:1.6; }
        .import-help-body code { display:block; background:#eef1ff; color:#333; padding:8px 10px; border-radius:6px; font-size:12px; margin-top:4px; white-space:pre-wrap; word-break:break-all; }
        .import-sample-btn { padding:8px 16px; font-size:13px; }
        .import-label { display:block; font-weight:bold; color:#667eea; margin:14px 0 6px; font-size:14px; }
        .import-textarea { width:100%; min-height:110px; padding:12px; border:2px solid #e0e0e0; border-radius:8px; font-size:13px; font-family:monospace; resize:vertical; }
        .import-textarea:focus { outline:none; border-color:#667eea; }
        .import-mode { margin:16px 0; display:flex; flex-direction:column; gap:8px; }
        .import-mode label { font-size:14px; color:#333; cursor:pointer; }
        .import-actions { display:flex; gap:12px; margin-top:8px; }
        .import-actions button { flex:1; padding:14px; font-size:15px; }
        .import-result { margin-top:16px; }
        .import-result__box { border-radius:8px; padding:16px; font-size:14px; line-height:1.7; }
        .import-result__ok { background:#eaf7ee; border-left:4px solid #28a745; color:#1e7e34; }
        .import-result__warn { background:#fff3cd; border-left:4px solid #ffc107; color:#856404; }
        .import-result__err { background:#fdecea; border-left:4px solid #dc3545; color:#c0392b; }
        .import-file { display:block; width:100%; padding:14px; border:2px dashed #667eea; border-radius:8px; background:#f8f9ff; font-size:14px; cursor:pointer; }
        .import-file-info { margin-top:8px; font-size:13px; color:#28a745; font-weight:bold; min-height:18px; }
        button { padding: 12px 24px; background: #667eea; color: white; border: none; border-radius: 6px; font-size: 14px; font-weight: bold; cursor: pointer; transition: background 0.3s; }
        button:hover { background: #764ba2; }
        .btn-secondary { background: #6c757d; }
        .btn-secondary:hover { background: #5a6268; }
        .dashboard-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 20px; }
        .stat-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px; text-align: center; }
        .stat-card h3 { font-size: 12px; opacity: 0.9; margin-bottom: 10px; text-transform: uppercase; }
        .stat-card .value { font-size: 32px; font-weight: bold; }
        .loading { text-align: center; padding: 40px; color: #666; }
        .loading::after { content: ''; display: inline-block; width: 20px; height: 20px; border: 3px solid #f3f3f3; border-top: 3px solid #667eea; border-radius: 50%; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

        /* ===== 出題画面 ===== */
        .quiz-top { display: flex; justify-content: flex-end; align-items: center; margin-bottom: 15px; }
        .quiz-abort { background: transparent; color: #888; border: 1px solid #ddd; padding: 6px 14px; font-size: 12px; font-weight: normal; }
        .quiz-abort:hover { background: #f5f5f5; color: #666; }
        .quiz-progress-bar { height: 8px; background: #eee; border-radius: 4px; overflow: hidden; margin-bottom: 8px; }
        .quiz-progress-bar__fill { height: 100%; background: linear-gradient(90deg, #667eea, #764ba2); transition: width 0.3s; }
        .quiz-progress-label { text-align: center; color: #667eea; font-weight: bold; margin-bottom: 15px; font-size: 15px; }
        .quiz-meta { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 15px; }
        .quiz-meta__item { font-size: 12px; color: #666; background: #f0f0f5; padding: 4px 10px; border-radius: 10px; }
        .quiz-meta__item--select-type { background: #e8eeff; color: #667eea; font-weight: bold; }
        .quiz-card { background: #f8f9ff; border-radius: 8px; padding: 20px; margin-bottom: 20px; border-left: 4px solid #667eea; }
        .quiz-question { font-size: 18px; line-height: 1.7; color: #333; margin-bottom: 20px; white-space: pre-wrap; }
        .quiz-choices { display: flex; flex-direction: column; gap: 12px; }
        .quiz-choice { display: flex; align-items: flex-start; gap: 12px; padding: 15px; border: 2px solid #e0e0e0; border-radius: 8px; cursor: pointer; transition: all 0.2s; background: white; }
        .quiz-choice:hover { border-color: #667eea; }
        .quiz-choice--selected { border-color: #667eea; background: #eef1ff; }
        .quiz-choice input { margin-top: 4px; width: 18px; height: 18px; cursor: pointer; }
        .quiz-choice span { flex: 1; line-height: 1.6; color: #333; }
        .quiz-next-btn { width: 100%; padding: 16px; font-size: 16px; }

        /* ===== ミックス出題の絞り込みフォーム ===== */
        .mix-form { max-width: 420px; margin: 0 auto 20px; text-align: left; }
        .mix-form label { display: block; font-weight: bold; color: #667eea; margin: 14px 0 6px; font-size: 14px; }
        .mix-select { width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 15px; background: white; cursor: pointer; }
        .mix-start-btn { width: 100%; padding: 14px; font-size: 16px; margin-top: 20px; }

        /* ===== 結果画面 ===== */
        .result-summary { text-align: center; background: #f8f9ff; border-radius: 8px; padding: 24px; margin-bottom: 24px; }
        .result-summary__label { color: #888; font-size: 14px; }
        .result-summary__score { font-size: 40px; font-weight: bold; color: #667eea; margin: 5px 0; }
        .result-summary__score span { font-size: 18px; color: #888; }
        .result-summary__rate { font-size: 18px; color: #333; }
        .result-summary__avg-time { color: #888; font-size: 13px; margin-top: 12px; }
        .result-summary__avg-time span { color: #667eea; font-weight: bold; margin-left: 6px; }
        .result-item { border: 1px solid #eee; border-left: 4px solid #ccc; border-radius: 8px; padding: 16px; margin-bottom: 14px; }
        .result-item--correct { border-left-color: #28a745; }
        .result-item--wrong { border-left-color: #dc3545; }
        .result-item__no { font-weight: bold; margin-bottom: 8px; font-size: 15px; }
        .result-item--correct .result-item__no { color: #28a745; }
        .result-item--wrong .result-item__no { color: #dc3545; }
        .result-item__meta { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px; }
        .result-item__meta-tag { font-size: 11px; color: #666; background: #f0f0f5; padding: 2px 8px; border-radius: 8px; }
        .result-item__question { font-size: 15px; line-height: 1.7; color: #333; margin-bottom: 12px; white-space: pre-wrap; }
        .result-item__answer-block { margin-bottom: 10px; }
        .result-item__answer-label { font-size: 12px; color: #888; margin-bottom: 4px; font-weight: bold; }
        .result-item__answer-list { list-style: none; padding-left: 0; }
        .result-item__answer-list li { padding: 6px 10px; border-radius: 5px; margin-bottom: 4px; font-size: 14px; background: #f5f5f5; }
        .result-item__answer-list--wrong li { background: #fdecea; color: #c0392b; }
        .result-item__answer-list--correct li { background: #eaf7ee; color: #1e7e34; }
        .result-item__explanation-label { font-size: 12px; color: #888; font-weight: bold; margin-bottom: 4px; }
        .result-item__explanation { font-size: 14px; line-height: 1.7; color: #555; margin-bottom: 8px; white-space: pre-wrap; }
        .result-item__link { color: #667eea; font-size: 13px; }
        .result-back-btn { width: 100%; padding: 16px; font-size: 16px; margin-top: 10px; }
        /* ===== 追加: ダッシュボード拡充 ===== */
        .dash-section { margin-bottom: 22px; }
        .dash-section__title { font-size: 14px; color: #667eea; margin-bottom: 10px; border-left: 4px solid #667eea; padding-left: 8px; }
        .dash-list { display: flex; flex-direction: column; gap: 12px; }
        .dash-row__head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; }
        .dash-row__name { font-size: 14px; color: #333; display: flex; align-items: center; gap: 6px; }
        .dash-row__val { font-size: 13px; color: #666; }
        .dash-bar { height: 10px; background: #eee; border-radius: 5px; overflow: hidden; }
        .dash-bar__fill { height: 100%; background: linear-gradient(90deg, #667eea, #764ba2); transition: width 0.4s; }
        .dash-bar__fill--warn { background: linear-gradient(90deg, #f0932b, #eb4d4b); }
        .dash-note { font-size: 12px; color: #888; margin: 4px 0 18px; display: flex; align-items: center; gap: 6px; }
        .warn-badge { display: inline-flex; align-items: center; justify-content: center; width: 18px; height: 18px; border-radius: 50%; background: #1f3a93; color: #fff; font-size: 12px; font-weight: bold; line-height: 1; }
        .dash-empty { color: #888; font-size: 13px; padding: 10px 0; }

        /* ===== 追加: 本番モードのタイマー・合否バッジ ===== */
        .quiz-top--timer { justify-content: space-between; }
        .quiz-timer { font-size: 15px; font-weight: bold; color: #667eea; background: #eef1ff; padding: 6px 14px; border-radius: 16px; }
        .quiz-timer--warn { color: #fff; background: #eb4d4b; }
        .pass-badge { display: inline-block; margin-top: 10px; padding: 8px 18px; border-radius: 20px; font-size: 14px; font-weight: bold; }
        .pass-badge--ok { background: #eaf7ee; color: #1e7e34; border: 2px solid #28a745; }
        .pass-badge--ng { background: #fdecea; color: #c0392b; border: 2px solid #dc3545; }

        /* ===== 追加: 受験者選択(プルダウン+追加) ===== */
        .add-student { margin-top: 26px; padding-top: 18px; border-top: 1px dashed #ddd; }
        .add-student__label { font-size: 12px; color: #888; margin-bottom: 8px; }
        .add-student__input { width: 100%; padding: 10px 12px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 14px; margin-bottom: 10px; }
        .add-student__msg { font-size: 13px; margin-top: 8px; min-height: 18px; }

        /* ===== 追加: 用語学習 ===== */
        .term-filter { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; justify-content: center; margin-bottom: 20px; }
        .term-filter label { font-size: 13px; color: #666; }
        .flash-card { background: #f8f9ff; border-left: 4px solid #667eea; border-radius: 8px; padding: 28px 20px; text-align: center; margin-bottom: 20px; }
        .flash-card__meta { display: flex; gap: 8px; justify-content: center; margin-bottom: 14px; flex-wrap: wrap; }
        .flash-card__term { font-size: 24px; font-weight: bold; color: #333; margin-bottom: 16px; line-height: 1.6; }
        .flash-card__definition { font-size: 15px; line-height: 1.9; color: #555; text-align: left; white-space: pre-wrap; }
        .flash-actions { display: flex; gap: 12px; }
        .flash-actions button { flex: 1; padding: 14px; }
        .btn-learned { background: #28a745; }
        .btn-learned:hover { background: #1e7e34; }
        .term-quiz-choice { display: block; width: 100%; text-align: left; padding: 14px; margin-bottom: 10px; background: #fff; border: 2px solid #e0e0e0; border-radius: 8px; color: #333; font-weight: normal; }
        .term-quiz-choice:hover { border-color: #667eea; background: #eef1ff; }
        .term-judge { text-align: center; font-size: 18px; font-weight: bold; margin: 14px 0; }
        .term-judge--ok { color: #1e7e34; }
        .term-judge--ng { color: #c0392b; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📋 模擬試験アプリ</h1>
        <div class="subtitle">マルチユーザー対応版</div>

        <!-- 受験者選択 -->
        <div id="nameSelectScreen" class="screen active">
            <div class="name-select-screen">
                <h2>受験者を選択してください</h2>
                <select id="studentSelect" class="mix-select" style="max-width:360px; margin:0 auto 14px;">
                    <option value="">読み込み中...</option>
                </select>
                <div><button type="button" onclick="selectStudentFromSelect()">この受験者で始める</button></div>

                <div class="add-student">
                    <p class="add-student__label">名簿にない場合はここで追加できます（同姓同名がいるときは「田中太郎(営業)」のように区別できる言葉を足してください）</p>
                    <input type="text" id="newStudentName" class="add-student__input" placeholder="氏名を入力">
                    <button type="button" class="btn-secondary" onclick="addStudent()">受験者を追加</button>
                    <p id="addStudentMsg" class="add-student__msg"></p>
                </div>
            </div>
        </div>

        <!-- メニュー -->
        <div id="menuScreen" class="screen">
            <div class="student-info">
                <p class="student-name" id="currentStudentName"></p>
                <p style="font-size: 12px; margin-top: 5px;"><a href="javascript:changeStudent()" style="color: #667eea; text-decoration: none;">別の学生に変更</a></p>
            </div>
            <h2 style="color: #667eea; margin-bottom: 20px; text-align: center;">何をしますか？</h2>
            <div class="menu-grid">
                <div class="menu-item" onclick="goToDashboard()"><h3>📊 ダッシュボード</h3><p>成績と履歴を確認</p></div>
                <div class="menu-item" onclick="startCategoryQuiz()"><h3>📝 カテゴリー別</h3><p>カテゴリーを選んで10問</p></div>
                <div class="menu-item" onclick="startTypeQuiz()"><h3>🎯 出題タイプ別</h3><p>出題タイプを選んで10問</p></div>
                <div class="menu-item" onclick="startMixQuiz()"><h3>🔀 ミックス出題</h3><p>カテゴリー×タイプで10問</p></div>
                <div class="menu-item" onclick="startRandomQuiz()"><h3>🎲 ランダム</h3><p>全問題からランダム10問</p></div>
                <div class="menu-item" onclick="startHonbanQuiz()"><h3>🏁 本番モード</h3><p>全問題から50問・60分</p></div>
                <div class="menu-item" onclick="startReviewQuiz()"><h3>🔁 復習モード</h3><p>間違えた問題から10問</p></div>
                <div class="menu-item" onclick="goToTermMenu()"><h3>📚 用語学習</h3><p>用語を覚える・用語クイズ</p></div>
                <div class="menu-item" onclick="showImport()"><h3>📥 問題インポート</h3><p>CSVで問題を追加・更新</p></div>
            </div>
      <p id="dataVersionNote" class="dash-note" style="margin-top:16px;"></p>
        </div>

        <!-- カテゴリー選択 -->
        <div id="categorySelectScreen" class="screen">
            <h2 style="color: #667eea; margin-bottom: 20px; text-align: center;">カテゴリーを選択してください</h2>
            <div class="menu-grid">
                <div class="menu-item" onclick="selectQuizCategory('アプリ')"><h3>アプリ</h3></div>
                <div class="menu-item" onclick="selectQuizCategory('スペース')"><h3>スペース</h3></div>
                <div class="menu-item" onclick="selectQuizCategory('管理者')"><h3>管理者</h3></div>
                <div class="menu-item" onclick="selectQuizCategory('その他')"><h3>その他</h3></div>
            </div>
            <button class="btn-secondary" onclick="goToMenu()">← メインメニューに戻る</button>
        </div>

        <!-- 出題タイプ選択 -->
        <div id="typeSelectScreen" class="screen">
            <h2 style="color: #667eea; margin-bottom: 20px; text-align: center;">出題タイプを選択してください</h2>
            <div class="menu-grid">
                <div class="menu-item" onclick="selectQuizType('知識問題')"><h3>知識問題</h3></div>
                <div class="menu-item" onclick="selectQuizType('実務問題')"><h3>実務問題</h3></div>
                <div class="menu-item" onclick="selectQuizType('アプリストア問題')"><h3>アプリストア問題</h3></div>
            </div>
            <button class="btn-secondary" onclick="goToMenu()">← メインメニューに戻る</button>
        </div>

        <!-- ミックス出題（絞り込み選択） -->
        <div id="mixSelectScreen" class="screen">
            <h2 style="color: #667eea; margin-bottom: 8px; text-align: center;">ミックス出題</h2>
            <p style="text-align:center; color:#888; font-size:13px; margin-bottom:20px;">カテゴリーと出題タイプで絞り込めます（両方「指定なし」でもOK）</p>
            <div class="mix-form">
                <label>カテゴリー</label>
                <select id="mixCategory" class="mix-select">
                    <option value="">指定なし（すべて）</option>
                    <option value="アプリ">アプリ</option>
                    <option value="スペース">スペース</option>
                    <option value="管理者">管理者</option>
                    <option value="その他">その他</option>
                </select>
                <label>出題タイプ</label>
                <select id="mixType" class="mix-select">
                    <option value="">指定なし（すべて）</option>
                    <option value="知識問題">知識問題</option>
                    <option value="実務問題">実務問題</option>
                    <option value="アプリストア問題">アプリストア問題</option>
                </select>
                <button type="button" class="mix-start-btn" onclick="startMix()">この条件で開始</button>
            </div>
            <button class="btn-secondary" onclick="goToMenu()">← メインメニューに戻る</button>
        </div>

        <!-- 出題 -->
        <div id="quizScreen" class="screen">
            <div id="quizContent"></div>
        </div>

        <!-- 結果 -->
        <div id="resultScreen" class="screen">
            <div id="resultContent"></div>
        </div>

        <!-- ダッシュボード -->
        <div id="dashboardScreen" class="screen">
            <h2 style="color: #667eea; margin-bottom: 20px; text-align: center;">📊 ダッシュボード</h2>
            <div class="dashboard-grid">
                <div class="stat-card">
                    <h3>受験回数</h3>
                    <div class="value" id="statExamCount">0</div>
                </div>
                <div class="stat-card">
                    <h3>平均正答率</h3>
                    <div class="value" id="statAvgRate">0<span style="font-size: 16px;">%</span></div>
                </div>
            </div>

            <div class="dash-section">
                <h3 class="dash-section__title">カテゴリー別の正答率</h3>
                <div id="dashCategory" class="dash-list"><div class="loading"></div></div>
            </div>

            <div class="dash-section">
                <h3 class="dash-section__title">出題タイプ別の正答率</h3>
                <div id="dashType" class="dash-list"></div>
            </div>

      <div class="dash-section">
        <h3 class="dash-section__title">用語学習</h3>
        <div id="dashTerm" class="dash-list"></div>
      </div>

            <p class="dash-note">正答率が70%未満の項目には <span class="warn-badge">!</span> が付きます</p>
            <button onclick="goToMenu()">← メインメニューに戻る</button>
        </div>

        <!-- 用語学習メニュー -->
        <div id="termMenuScreen" class="screen">
            <!-- 受験者の表示と切り替え（誤操作を防ぐため、学習中の画面には置かない） -->
            <div class="student-info">
                <p class="student-name" id="currentStudentNameTerm"></p>
                <p style="font-size: 12px; margin-top: 5px;"><a href="javascript:changeStudent()" style="color: #667eea; text-decoration: none;">別の学生に変更</a></p>
            </div>
            <h2 style="color: #667eea; margin-bottom: 16px; text-align: center;">📚 用語学習</h2>

            <div class="term-filter">
                <label>章</label>
                <select id="termChapter" class="mix-select" style="width:auto;" onchange="loadTermDashboard()">
                    <option value="">すべて</option>
                    <option value="1">1章</option>
                    <option value="2">2章</option>
                    <option value="3">3章</option>
                    <option value="4">4章</option>
                    <option value="5">5章</option>
                    <option value="6">6章</option>
                </select>
                <label>難易度</label>
                <select id="termLevel" class="mix-select" style="width:auto;" onchange="loadTermDashboard()">
                    <option value="">すべて</option>
                    <option value="基礎">基礎</option>
                    <option value="応用">応用</option>
                </select>
            </div>

            <div class="dashboard-grid">
                <div class="stat-card">
                    <h3>覚えた用語</h3>
                    <div class="value" id="termLearnedCount">0</div>
                </div>
                <div class="stat-card">
                    <h3>達成率</h3>
                    <div class="value" id="termRate">0<span style="font-size: 16px;">%</span></div>
                </div>
            </div>

            <div class="dash-section">
                <h3 class="dash-section__title">章別の達成率</h3>
                <div id="termByChapter" class="dash-list"></div>
            </div>
            <div class="dash-section">
                <h3 class="dash-section__title">難易度別の達成率</h3>
                <div id="termByLevel" class="dash-list"></div>
            </div>
            <div class="dash-section">
                <h3 class="dash-section__title">用語クイズの正答率（章別）</h3>
                <div id="termQuizByChapter" class="dash-list"></div>
            </div>
            <div class="dash-section">
                <h3 class="dash-section__title">用語クイズの正答率（難易度別）</h3>
                <div id="termQuizByLevel" class="dash-list"></div>
            </div>
            <p class="dash-note">正答率・達成率が70%未満の項目には <span class="warn-badge">!</span> が付きます</p>

            <div class="menu-grid">
                <div class="menu-item" onclick="startFlashcard()"><h3>🃏 フラッシュカード</h3><p>用語と意味を1枚ずつ確認</p></div>
                <div class="menu-item" onclick="startTermQuiz()"><h3>❓ 用語クイズ</h3><p>覚えた用語から4択で出題</p></div>
            </div>
            <button class="btn-secondary" onclick="goToMenu()">← メインメニューに戻る</button>
        </div>

        <!-- フラッシュカード -->
        <div id="flashScreen" class="screen">
            <div id="flashContent"></div>
        </div>

        <!-- 用語クイズ -->
        <div id="termQuizScreen" class="screen">
            <div id="termQuizContent"></div>
        </div>

        <!-- 問題インポート -->
        <div id="importScreen" class="screen">
            <h2 style="color: #667eea; margin-bottom: 8px; text-align: center;">📥 問題インポート</h2>
            <p style="text-align:center; color:#888; font-size:13px; margin-bottom:16px;">CSVファイルを選ぶだけ。形式は自動で判別します（複数ファイルまとめてOK）</p>

            <label class="import-label">CSVファイルを選ぶ（選ぶとファイル名が下に出ます）</label>
            <input type="file" id="importFile" class="import-file" accept=".csv,text/csv" multiple onchange="loadImportFiles(this)">
            <div id="importFileInfo" class="import-file-info"></div>

            <div class="import-mode">
                <label><input type="radio" name="importMode" value="append" checked> 追記（同じIDは上書き・他はそのまま）※おすすめ</label>
                <label><input type="radio" name="importMode" value="replace"> 全置換（今の問題を全部消して入れ替え）</label>
            </div>

            <div class="import-actions">
                <button type="button" class="btn-secondary" onclick="previewImport()">チェック（プレビュー）</button>
                <button type="button" onclick="runImport()">取り込む</button>
            </div>
            <div id="importResult" class="import-result"></div>

            <button class="btn-secondary" onclick="goToMenu()" style="margin-top:16px;">← メインメニューに戻る</button>
        </div>
    </div>

    <script>
        const STORAGE_KEY_STUDENT = 'kintone_exam_student_name';
        const QUIZ_COUNT = 10; // カテゴリー別の出題数

        let currentStudent = null;
        // 受験者を選び直したあと、どの画面へ戻るか（用語学習から変更した場合は用語メニューへ戻す）
        let studentReturnScreen = 'menuScreen';

        // 出題用の状態（キントーン版 state に相当）
        let quizQuestions = [];
        let answers = [];
        let currentQuizIndex = 0;
        let quizStartTime = 0;
        let selectedCategory = '';
        let selectedType = '';
        let examMode = '';
        let quizLabel = '';

        // 本番モード用のタイマー状態
        const HONBAN_COUNT = 50;          // 本番モードの出題数
        const HONBAN_TIME_LIMIT = 60 * 60; // 本番モードの制限時間(秒)
        const PASS_LINE = 70;             // 合格ライン(%)
        let quizTimeLimit = 0;
        let quizTimerId = null;
        let examFinished = false;

        // 用語学習用の状態
        let termAll = [];
        let termLearnedIds = [];
        let flashList = [];
        let flashIndex = 0;
        let flashRevealed = false;
        let termQuizList = [];
        let termQuizIndex = 0;
        let termQuizCorrect = 0;
        let termQuizAnswered = false;

        function init() {
            const saved = localStorage.getItem(STORAGE_KEY_STUDENT);
            if (saved) {
                currentStudent = saved;
                showScreen('menuScreen');
                updateStudentInfo();
            } else {
                loadStudents();
            }
        }

        /* ===== GAS 呼び出し（google.script.run を Promise 化） ===== */
        function gasRun(fnName, args) {
            return new Promise(function (resolve, reject) {
                var runner = google.script.run
                    .withSuccessHandler(function (res) { resolve(res); })
                    .withFailureHandler(function (err) {
                        console.error('GAS Error:', err);
                        alert('データ取得失敗: ' + (err && err.message ? err.message : err));
                        reject(err);
                    });
                runner[fnName].apply(runner, args || []);
            });
        }

        /* ===== 受験者選択 ===== */
        function loadStudents(keepName) {
            return gasRun('getStudents', []).then(function (students) {
                students = students || [];
                const sel = document.getElementById('studentSelect');
                sel.innerHTML = '';
                if (students.length === 0) {
                    const op = document.createElement('option');
                    op.value = '';
                    op.textContent = '（受験者が登録されていません）';
                    sel.appendChild(op);
                    return;
                }
                students.forEach(function (student) {
                    const op = document.createElement('option');
                    op.value = student.student_name;
                    op.textContent = student.student_name;
                    sel.appendChild(op);
                });
                if (keepName) sel.value = keepName;
            });
        }
        function selectStudentFromSelect() {
            const sel = document.getElementById('studentSelect');
            const name = sel ? sel.value : '';
            if (!name) {
                alert('受験者を選んでください');
                return;
            }
            selectStudent(name);
        }
        function addStudent() {
            const input = document.getElementById('newStudentName');
            const msg = document.getElementById('addStudentMsg');
            const name = (input.value || '').trim();
            msg.style.color = '#c33';
            if (!name) {
                msg.textContent = '氏名を入力してください';
                return;
            }
            msg.style.color = '#888';
            msg.textContent = '登録しています...';
            gasRun('addStudent', [name]).then(function (res) {
                if (res && res.ok) {
                    msg.style.color = '#1e7e34';
                    msg.textContent = '「' + name + '」を追加しました';
                    input.value = '';
                    loadStudents(name);
                } else {
                    msg.style.color = '#c33';
                    msg.textContent = (res && res.message) ? res.message : '追加できませんでした';
                }
            });
        }
        function selectStudent(name) {
            currentStudent = name;
            localStorage.setItem(STORAGE_KEY_STUDENT, name);
            const back = studentReturnScreen;
            studentReturnScreen = 'menuScreen';
            if (back === 'termMenuScreen') {
                showScreen('termMenuScreen');
                updateStudentInfo();
                loadTermDashboard();
            } else {
                showScreen('menuScreen');
                updateStudentInfo();
            }
        }
        function changeStudent() {
            // 変更を始めた画面を覚えておき、選択後にそこへ戻す
            const active = document.querySelector('.screen.active');
            studentReturnScreen = (active && active.id === 'termMenuScreen') ? 'termMenuScreen' : 'menuScreen';
            localStorage.removeItem(STORAGE_KEY_STUDENT);
            currentStudent = null;
            showScreen('nameSelectScreen');
            loadStudents();
        }
        function updateStudentInfo() {
            document.getElementById('currentStudentName').textContent = currentStudent;
            var termNameEl = document.getElementById('currentStudentNameTerm');
            if (termNameEl) termNameEl.textContent = currentStudent;
        }

        /* ===== 画面切り替え ===== */
        function showScreen(screenId) {
      if (screenId === "menuScreen") { setTimeout(loadDataVersion, 0); }
            document.querySelectorAll('.screen').forEach(function (s) { s.classList.remove('active'); });
            document.getElementById(screenId).classList.add('active');
        }
        function goToMenu() {
      loadDataVersion();
            showScreen('menuScreen');
            window.scrollTo(0, 0);
        }
        function goToDashboard() {
            showScreen('dashboardScreen');
            loadDashboard();
        }
        // メニュー下部に知識問題データのバージョンを表示
    function loadDataVersion() {
      var el = document.getElementById("dataVersionNote");
      if (!el) return;
      gasRun("getDataVersion", []).then(function (v) {
        if (!v) { el.textContent = ""; return; }
        var parts = [];
        parts.push("知識問題データ　バージョン " + v.version);
        if (v.target_exam) parts.push("対応試験：" + v.target_exam);
        if (v.imported_count) parts.push(v.imported_count + "問");
        if (v.imported_at) parts.push(v.imported_at + " 取込");
        el.textContent = parts.join("　/　");
      });
    }

    // ダッシュボードに用語学習のサマリーを描画
    function loadDashboardTerm() {
      var el = document.getElementById("dashTerm");
      if (!el) return;
      el.innerHTML = '<div class="loading"></div>';
      gasRun("getTermDashboard", [currentStudent, "", ""]).then(function (t) {
        if (!t) { renderRateList("dashTerm", [], "用語データがありません"); return; }
        var qt = 0, qd = 0;
        (t.quiz_by_chapter || []).forEach(function (r) { qt += (r.total || 0); qd += (r.done || 0); });
        var rows = [
          { key: "覚えた用語", total: t.total || 0, done: t.learned || 0 },
          { key: "用語クイズの正答率", total: qt, done: qd }
        ];
        renderRateList("dashTerm", rows, "用語データがありません");
      });
    }

    function loadDashboard() {
      loadDashboardTerm();
            document.getElementById('dashCategory').innerHTML = '<div class="loading"></div>';
            document.getElementById('dashType').innerHTML = '';
            gasRun('getDashboardStats', [currentStudent]).then(function (st) {
                st = st || { exam_count: 0, avg_rate: 0, by_category: [], by_type: [] };
                document.getElementById('statExamCount').textContent = st.exam_count || 0;
                document.getElementById('statAvgRate').innerHTML = (st.avg_rate || 0) + '<span style="font-size: 16px;">%</span>';
                renderRateList('dashCategory', st.by_category, 'まだ回答したデータがありません');
                renderRateList('dashType', st.by_type, 'まだ回答したデータがありません');
            });
        }

        // 「項目名 / 正答率 / バー / 70%未満は警告バッジ」の共通描画
        function renderRateList(elementId, rows, emptyMessage) {
            const el = document.getElementById(elementId);
            if (!el) return;
            rows = rows || [];
            if (rows.length === 0) {
                el.innerHTML = '<p class="dash-empty">' + escapeHtml(emptyMessage || 'データがありません') + '</p>';
                return;
            }
            let html = '';
            rows.forEach(function (r) {
                const total = r.total || 0;
                const done = r.done || 0;
                const rate = total > 0 ? Math.round((done / total) * 100) : 0;
                const warn = (total > 0 && rate < PASS_LINE);
                html += '<div class="dash-row">';
                html += '  <div class="dash-row__head">';
                html += '    <span class="dash-row__name">' + escapeHtml(r.key) + (warn ? '<span class="warn-badge">!</span>' : '') + '</span>';
                html += '    <span class="dash-row__val">' + (total > 0 ? rate + '%（' + done + '/' + total + '）' : '未実施') + '</span>';
                html += '  </div>';
                html += '  <div class="dash-bar"><div class="dash-bar__fill' + (warn ? ' dash-bar__fill--warn' : '') + '" style="width:' + rate + '%"></div></div>';
                html += '</div>';
            });
            el.innerHTML = html;
        }

        /* ===== 共通ユーティリティ ===== */
        function shuffleArr(arr) {
            for (let i = arr.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                const tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
            }
            return arr;
        }
        function isTrue(v) {
            if (v === true) return true;
            if (typeof v === 'number') return v === 1;
            const s = String(v).trim().toLowerCase();
            return s === 'true' || s === '1' || s === '○' || s === '◯' || s === '✓' || s === 'yes' || s === '正解';
        }
        function arraysEqual(a, b) {
            if (a.length !== b.length) return false;
            for (let i = 0; i < a.length; i++) { if (a[i] !== b[i]) return false; }
            return true;
        }
        function escapeHtml(s) {
            if (s == null) return '';
            return String(s)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        }

        /* ===== 出題（カテゴリー別・出題タイプ別 共通／キントーン版と同じ基本動作） ===== */
        function startCategoryQuiz() {
            showScreen('categorySelectScreen');
            window.scrollTo(0, 0);
        }
        function startTypeQuiz() {
            showScreen('typeSelectScreen');
            window.scrollTo(0, 0);
        }
        function selectQuizCategory(category) {
            loadQuiz('問題カテゴリー別', category, '', category);
        }
        function selectQuizType(type) {
            loadQuiz('出題タイプ別', '', type, type);
        }
        function startMixQuiz() {
            showScreen('mixSelectScreen');
            window.scrollTo(0, 0);
        }
        function startMix() {
            var cat = document.getElementById('mixCategory').value;
            var typ = document.getElementById('mixType').value;
            var parts = [];
            if (cat) parts.push(cat);
            if (typ) parts.push(typ);
            var label = parts.length ? parts.join('・') : '全問題';
            loadQuiz('複合', cat, typ, label);
        }
        function startRandomQuiz() {
            loadQuiz('ランダム', '', '', 'ランダム10問');
        }

        // 本番モード：全問題から50問・60分カウントダウン
        function startHonbanQuiz() {
            if (!confirm('本番モードを始めます。\\n全問題から50問・制限時間60分です。よろしいですか？')) return;
            loadQuiz('本番モード', '', '', '本番モード（50問・60分）', HONBAN_COUNT, HONBAN_TIME_LIMIT, null);
        }

        // 復習モード：過去に一度でも不正解だった問題から10問
        function startReviewQuiz() {
            gasRun('getWrongQuestionIds', [currentStudent]).then(function (ids) {
                ids = ids || [];
                if (ids.length === 0) {
                    alert('復習できる問題がありません。\\n（過去に不正解になった問題がまだないようです）');
                    return;
                }
                loadQuiz('復習モード', '', '', '復習モード（間違えた問題）', QUIZ_COUNT, 0, ids);
            }).catch(function () { /* 取得失敗時は gasRun 側でメッセージ表示済み */ });
        }

        // mode:履歴に残すモード名 / category,type:絞り込み条件 / label:画面表示用
        // count:出題数(省略時10) / timeLimitSec:制限時間(0なら無制限) / idFilter:出題対象の問題ID配列
        function loadQuiz(mode, category, type, label, count, timeLimitSec, idFilter) {
            examMode = mode;
            selectedCategory = category;
            selectedType = type;
            quizLabel = label;
            currentQuizIndex = 0;
            answers = [];
            quizQuestions = [];
            examFinished = false;
            quizTimeLimit = timeLimitSec || 0;
            clearQuizTimer();
            showScreen('quizScreen');
            window.scrollTo(0, 0);
            document.getElementById('quizContent').innerHTML =
                '<div class="loading"></div><p style="text-align:center;color:#666;">問題を準備しています...</p>';

            const quizCount = count || QUIZ_COUNT;

            gasRun('getQuestions', [category || null, type || null]).then(function (qs) {
                qs = qs || [];
                if (idFilter && idFilter.length) {
                    const idSet = {};
                    idFilter.forEach(function (id) { idSet[String(id)] = true; });
                    qs = qs.filter(function (q) { return idSet[String(q.question_id)]; });
                }
                if (qs.length === 0) {
                    document.getElementById('quizContent').innerHTML =
                        '<p style="text-align:center;color:#c33;margin-bottom:20px;">「' + escapeHtml(label) + '」の問題がまだありません。</p>' +
                        '<button class="btn-secondary" onclick="goToMenu()">← メインメニューに戻る</button>';
                    return null;
                }
                shuffleArr(qs);
                qs = qs.slice(0, quizCount);
                const qids = qs.map(function (q) { return q.question_id; });
                return gasRun('getChoicesBatch', [qids]).then(function (map) {
                    map = map || {};
                    qs.forEach(function (q) {
                        const choices = map[String(q.question_id)] || [];
                        shuffleArr(choices);
                        q.choices = choices;
                    });
                    return qs;
                });
            }).then(function (list) {
                if (!list) return;
                quizQuestions = list;
                answers = list.map(function () { return []; });
                quizStartTime = Date.now();
                renderQuiz();
                startQuizTimer();
                window.scrollTo(0, 0);
            });
        }

        /* ===== 本番モードのカウントダウンタイマー ===== */
        function startQuizTimer() {
            clearQuizTimer();
            if (!quizTimeLimit) return;
            updateQuizTimer();
            quizTimerId = setInterval(updateQuizTimer, 1000);
        }
        function clearQuizTimer() {
            if (quizTimerId) {
                clearInterval(quizTimerId);
                quizTimerId = null;
            }
        }
        function remainingSeconds() {
            const rest = quizTimeLimit - Math.floor((Date.now() - quizStartTime) / 1000);
            return rest > 0 ? rest : 0;
        }
        function updateQuizTimer() {
            if (!quizTimeLimit) return;
            const rest = remainingSeconds();
            const el = document.getElementById('quizTimer');
            if (el) {
                const m = Math.floor(rest / 60);
                const sec = rest % 60;
                el.textContent = '残り ' + m + ':' + (sec < 10 ? '0' + sec : sec);
                if (rest <= 300) el.classList.add('quiz-timer--warn');
                else el.classList.remove('quiz-timer--warn');
            }
            if (rest <= 0) {
                clearQuizTimer();
                alert('制限時間になりました。ここまでの回答で採点します。');
                finishExam();
            }
        }

        function renderQuiz() {
            const q = quizQuestions[currentQuizIndex];
            const total = quizQuestions.length;
            const idx = currentQuizIndex;
            const isMulti = q.select_type === '複数選択';
            const selected = answers[idx] || [];

            let html = '';
            html += '<div class="quiz-top' + (quizTimeLimit ? ' quiz-top--timer' : '') + '">';
            if (quizTimeLimit) {
                html += '  <span class="quiz-timer" id="quizTimer">残り --:--</span>';
            }
            html += '  <button type="button" class="quiz-abort" onclick="abortQuiz()">中断してメインメニューに戻る</button>';
            html += '</div>';
            html += '<div class="quiz-progress-bar"><div class="quiz-progress-bar__fill" style="width:' + Math.round((idx / total) * 100) + '%"></div></div>';
            html += '<p class="quiz-progress-label">第' + (idx + 1) + '問 / 全' + total + '問</p>';
            html += '<div class="quiz-meta">';
            html += '  <span class="quiz-meta__item">ID: ' + escapeHtml(q.question_id) + '</span>';
            html += '  <span class="quiz-meta__item">' + escapeHtml(q.question_category || '') + '</span>';
            html += '  <span class="quiz-meta__item">' + escapeHtml(q.question_type || '') + '</span>';
            html += '  <span class="quiz-meta__item quiz-meta__item--select-type">' + escapeHtml(q.select_type || '') + '</span>';
            html += '</div>';
            html += '<div class="quiz-card">';
            html += '  <p class="quiz-question">' + escapeHtml(q.question) + '</p>';
            html += '  <div class="quiz-choices">';
            q.choices.forEach(function (c, cIdx) {
                const isChecked = selected.indexOf(cIdx) !== -1;
                html += '<label class="quiz-choice' + (isChecked ? ' quiz-choice--selected' : '') + '">';
                html += '<input type="' + (isMulti ? 'checkbox' : 'radio') + '" name="quiz-choice" data-idx="' + cIdx + '"' + (isChecked ? ' checked' : '') + '>';
                html += '<span>' + escapeHtml(c.choice_text) + '</span>';
                html += '</label>';
            });
            html += '  </div>';
            html += '</div>';
            html += '<button type="button" class="quiz-next-btn" onclick="onNextClicked()">' + (idx === total - 1 ? '採点する' : '次の問題へ') + '</button>';

            document.getElementById('quizContent').innerHTML = html;
            bindQuizEvents();
            updateQuizTimer();
        }

        function bindQuizEvents() {
            const q = quizQuestions[currentQuizIndex];
            const isMulti = q.select_type === '複数選択';
            const inputs = document.querySelectorAll('#quizContent input[name="quiz-choice"]');
            inputs.forEach(function (input) {
                input.addEventListener('change', function () {
                    const i = parseInt(input.getAttribute('data-idx'), 10);
                    if (isMulti) {
                        const cur = answers[currentQuizIndex];
                        const pos = cur.indexOf(i);
                        if (input.checked && pos === -1) cur.push(i);
                        else if (!input.checked && pos !== -1) cur.splice(pos, 1);
                    } else {
                        answers[currentQuizIndex] = [i];
                    }
                    updateChoiceHighlight();
                });
            });
        }

        function updateChoiceHighlight() {
            const selected = answers[currentQuizIndex];
            const labels = document.querySelectorAll('#quizContent .quiz-choice');
            labels.forEach(function (label, idx) {
                if (selected.indexOf(idx) !== -1) label.classList.add('quiz-choice--selected');
                else label.classList.remove('quiz-choice--selected');
            });
        }

        function onNextClicked() {
            if (answers[currentQuizIndex].length === 0) {
                alert('選択肢を選んでください');
                return;
            }
            if (currentQuizIndex < quizQuestions.length - 1) {
                currentQuizIndex += 1;
                renderQuiz();
                window.scrollTo(0, 0);
            } else {
                finishExam();
            }
        }

        function abortQuiz() {
            if (confirm('中断すると今回の結果は記録されません。メニューに戻りますか？')) {
                clearQuizTimer();
                quizTimeLimit = 0;
                goToMenu();
            }
        }

        /* ===== 採点・履歴登録・結果画面 ===== */
        function finishExam() {
            if (examFinished) return;   // 時間切れと「採点する」の二重実行を防ぐ
            examFinished = true;
            clearQuizTimer();
            const durationSeconds = Math.round((Date.now() - quizStartTime) / 1000);
            const results = [];
            const displayInfo = [];
            let correctCount = 0;

            quizQuestions.forEach(function (q, idx) {
                const choices = q.choices;
                const correctIdxList = [];
                choices.forEach(function (c, cIdx) { if (isTrue(c.is_correct)) correctIdxList.push(cIdx); });

                const selectedIdx = answers[idx].slice().sort(function (a, b) { return a - b; });
                const correctSorted = correctIdxList.slice().sort(function (a, b) { return a - b; });
                const isCorrect = arraysEqual(selectedIdx, correctSorted);
                if (isCorrect) correctCount += 1;

                const selectedTextList = selectedIdx.map(function (cIdx) { return choices[cIdx].choice_text; });
                const correctTextList = correctIdxList.map(function (cIdx) { return choices[cIdx].choice_text; });
                const firstCorrect = choices[correctIdxList[0]];

                results.push({
                    question_id: q.question_id,
                    question_text: q.question,
                    selected_answer: selectedTextList.join('、'),
                    judgment: isCorrect ? '正解' : '不正解',
                    explanation_text: firstCorrect ? (firstCorrect.explanation_text || '') : '',
                    explanation_url: firstCorrect ? (firstCorrect.explanation_url || '') : ''
                });

                displayInfo.push({
                    selectedTextList: selectedTextList,
                    correctTextList: correctTextList,
                    isCorrect: isCorrect,
                    category: q.question_category,
                    type: q.question_type,
                    questionId: q.question_id,
                    questionText: q.question,
                    explanationText: firstCorrect ? (firstCorrect.explanation_text || '') : '',
                    explanationUrl: firstCorrect ? (firstCorrect.explanation_url || '') : ''
                });
            });

            document.getElementById('quizContent').innerHTML =
                '<div class="loading"></div><p style="text-align:center;color:#666;">結果を登録しています...</p>';

            const examData = {
                exam_mode: examMode,
                category: selectedCategory,
                type: selectedType,
                results: results,
                duration_seconds: durationSeconds
            };

            gasRun('postExamResult', [currentStudent, examData])
                .then(function () { renderResultScreen(correctCount, quizQuestions.length, durationSeconds, displayInfo); })
                .catch(function () { renderResultScreen(correctCount, quizQuestions.length, durationSeconds, displayInfo); });
        }

        function renderResultScreen(correctCount, total, durationSeconds, displayInfo) {
            const rate = Math.round((correctCount / total) * 1000) / 10;
            const avgSeconds = total > 0 ? Math.round((durationSeconds / total) * 10) / 10 : 0;

            let html = '';
            html += '<button type="button" class="result-back-btn" onclick="goToMenu()" style="margin-top:0; margin-bottom:16px;">← メインメニューに戻る</button>';
      html += '<div class="result-summary">';
            html += '  <p class="result-summary__label">結果（' + escapeHtml(quizLabel) + '）</p>';
            html += '  <p class="result-summary__score">' + correctCount + ' <span>/ ' + total + '問</span></p>';
            html += '  <p class="result-summary__rate">正答率 ' + rate + '%</p>';
            html += '  <p class="result-summary__avg-time">1問に要した平均時間<span>' + avgSeconds + '秒</span></p>';
            if (examMode === '本番モード') {
                const passed = rate >= PASS_LINE;
                html += '  <div class="pass-badge ' + (passed ? 'pass-badge--ok' : 'pass-badge--ng') + '">' +
                        (passed ? '合格ライン（' + PASS_LINE + '%）クリア' : '合格ライン（' + PASS_LINE + '%）未達') + '</div>';
            }
            html += '</div>';

            html += '<div class="result-list">';
            displayInfo.forEach(function (info, idx) {
                const isCorrect = info.isCorrect;
                html += '<div class="result-item result-item--' + (isCorrect ? 'correct' : 'wrong') + '">';
                html += '  <p class="result-item__no">第' + (idx + 1) + '問 ' + (isCorrect ? '正解' : '不正解') + '</p>';
                html += '  <div class="result-item__meta">';
                html += '    <span class="result-item__meta-tag">ID: ' + escapeHtml(info.questionId) + '</span>';
                html += '    <span class="result-item__meta-tag">' + escapeHtml(info.category || '') + '</span>';
                html += '    <span class="result-item__meta-tag">' + escapeHtml(info.type || '') + '</span>';
                html += '  </div>';
                html += '  <p class="result-item__question">' + escapeHtml(info.questionText) + '</p>';

                html += '  <div class="result-item__answer-block">';
                html += '    <p class="result-item__answer-label">あなたの回答</p>';
                html += '    <ul class="result-item__answer-list' + (isCorrect ? '' : ' result-item__answer-list--wrong') + '">';
                if (info.selectedTextList.length === 0) {
                    html += '<li>（未回答）</li>';
                } else {
                    info.selectedTextList.forEach(function (t) { html += '<li>' + escapeHtml(t) + '</li>'; });
                }
                html += '    </ul>';
                html += '  </div>';

                if (!isCorrect) {
                    html += '  <div class="result-item__answer-block">';
                    html += '    <p class="result-item__answer-label">正解</p>';
                    html += '    <ul class="result-item__answer-list result-item__answer-list--correct">';
                    info.correctTextList.forEach(function (t) { html += '<li>' + escapeHtml(t) + '</li>'; });
                    html += '    </ul>';
                    html += '  </div>';
                }

                if (info.explanationText) {
                    html += '  <p class="result-item__explanation-label">解説</p>';
                    html += '  <p class="result-item__explanation">' + escapeHtml(info.explanationText) + '</p>';
                }
                if (info.explanationUrl) {
                    html += '  <a class="result-item__link" href="' + encodeURI(info.explanationUrl) + '" target="_blank" rel="noopener">参考URLを開く</a>';
                }
                html += '</div>';
            });
            html += '</div>';

            html += '<button type="button" class="result-back-btn" onclick="goToMenu()">← メインメニューに戻る</button>';

            document.getElementById('resultContent').innerHTML = html;
            showScreen('resultScreen');
            window.scrollTo(0, 0);
        }

        /* ===== 用語学習 ===== */
        function goToTermMenu() {
            showScreen('termMenuScreen');
            window.scrollTo(0, 0);
            updateStudentInfo();
            loadTermDashboard();
        }

        function getTermFilter() {
            var chEl = document.getElementById('termChapter');
            var lvEl = document.getElementById('termLevel');
            return {
                chapter: chEl ? chEl.value : '',
                level: lvEl ? lvEl.value : ''
            };
        }

        function loadTermDashboard() {
            var f = getTermFilter();
            document.getElementById('termByChapter').innerHTML = '<div class="loading"></div>';
            document.getElementById('termByLevel').innerHTML = '';
            document.getElementById('termQuizByChapter').innerHTML = '';
            document.getElementById('termQuizByLevel').innerHTML = '';

            gasRun('getTermDashboard', [currentStudent, f.chapter || null, f.level || null]).then(function (d) {
                d = d || {};
                document.getElementById('termLearnedCount').textContent = (d.learned || 0) + ' / ' + (d.total || 0);
                document.getElementById('termRate').innerHTML = (d.rate || 0) + '<span style="font-size: 16px;">%</span>';
                renderRateList('termByChapter', d.by_chapter, '用語データがまだありません');
                renderRateList('termByLevel', d.by_level, '用語データがまだありません');
                renderRateList('termQuizByChapter', d.quiz_by_chapter, 'まだ用語クイズの記録がありません');
                renderRateList('termQuizByLevel', d.quiz_by_level, 'まだ用語クイズの記録がありません');
            });
        }

        // 絞り込み条件に合う用語だけを返す
        function filterTerms(terms) {
            var f = getTermFilter();
            return (terms || []).filter(function (t) {
                if (f.chapter && String(t.chapter).split('章')[0] !== String(f.chapter)) return false;
                if (f.level && String(t.level) !== String(f.level)) return false;
                return true;
            });
        }

        /* ----- フラッシュカード ----- */
        function startFlashcard() {
            showScreen('flashScreen');
            window.scrollTo(0, 0);
            document.getElementById('flashContent').innerHTML =
                '<div class="loading"></div><p style="text-align:center;color:#666;">用語を準備しています...</p>';

            Promise.all([gasRun('getTerms', [null]), gasRun('getTermProgress', [currentStudent])])
                .then(function (res) {
                    termAll = res[0] || [];
                    var prog = res[1] || {};
                    termLearnedIds = prog.learned_term_ids || [];

                    var list = filterTerms(termAll);
                    if (list.length === 0) {
                        document.getElementById('flashContent').innerHTML =
                            '<p style="text-align:center;color:#c33;margin-bottom:20px;">この条件の用語がありません。</p>' +
                            '<button class="btn-secondary" onclick="goToTermMenu()">← 用語学習に戻る</button>';
                        return;
                    }
                    shuffleArr(list);
                    flashList = list;
                    flashIndex = 0;
                    flashRevealed = false;
                    renderFlash();
                });
        }

        function renderFlash() {
            if (flashIndex >= flashList.length) {
                document.getElementById('flashContent').innerHTML =
                    '<p style="text-align:center;color:#667eea;font-size:18px;font-weight:bold;margin-bottom:20px;">この条件の用語をひと通り確認しました。</p>' +
                    '<button onclick="goToTermMenu()">← 用語学習に戻る</button>';
                return;
            }
            var t = flashList[flashIndex];
            var learned = termLearnedIds.indexOf(t.term_id) !== -1;

            var html = '';
            html += '<div class="quiz-top">';
            html += '  <button type="button" class="quiz-abort" onclick="goToTermMenu()">やめて用語学習メニューへ戻る</button>';
            html += '</div>';
            html += '<div class="quiz-progress-bar"><div class="quiz-progress-bar__fill" style="width:' + Math.round((flashIndex / flashList.length) * 100) + '%"></div></div>';
            html += '<p class="quiz-progress-label">' + (flashIndex + 1) + ' / ' + flashList.length + '語</p>';

            html += '<div class="flash-card">';
            html += '  <div class="flash-card__meta">';
            html += '    <span class="quiz-meta__item">' + escapeHtml(t.term_id || '') + '</span>';
            html += '    <span class="quiz-meta__item">' + escapeHtml(String(t.chapter || '')) + '</span>';
            html += '    <span class="quiz-meta__item quiz-meta__item--select-type">' + escapeHtml(t.level || '') + '</span>';
            if (learned) html += '    <span class="quiz-meta__item" style="background:#eaf7ee;color:#1e7e34;">覚えた</span>';
            html += '  </div>';
            html += '  <p class="flash-card__term">' + escapeHtml(t.term) + '</p>';
            if (flashRevealed) {
                html += '  <p class="flash-card__definition">' + escapeHtml(t.definition || '') + '</p>';
                if (t.explanation_url) {
                    html += '  <p style="margin-top:12px;text-align:left;"><a class="result-item__link" href="' + encodeURI(t.explanation_url) + '" target="_blank" rel="noopener">参考URLを開く</a></p>';
                }
            }
            html += '</div>';

            if (!flashRevealed) {
                html += '<button type="button" class="quiz-next-btn" onclick="flashReveal()">意味を見る</button>';
            } else {
                html += '<div class="flash-actions">';
                html += '  <button type="button" class="btn-secondary" onclick="flashNext(false)">まだ</button>';
                html += '  <button type="button" class="btn-learned" onclick="flashNext(true)">覚えた</button>';
                html += '</div>';
            }

            document.getElementById('flashContent').innerHTML = html;
        }

        function flashReveal() {
            flashRevealed = true;
            renderFlash();
        }

        function flashNext(learned) {
            var t = flashList[flashIndex];
            if (learned && t && termLearnedIds.indexOf(t.term_id) === -1) {
                termLearnedIds.push(t.term_id);
                gasRun('postTermLearned', [currentStudent, t.term_id]);  // 即時記録
            }
            flashIndex += 1;
            flashRevealed = false;
            renderFlash();
            window.scrollTo(0, 0);
        }

        /* ----- 用語クイズ（覚えた用語から4択） ----- */
        function startTermQuiz() {
            showScreen('termQuizScreen');
            window.scrollTo(0, 0);
            document.getElementById('termQuizContent').innerHTML =
                '<div class="loading"></div><p style="text-align:center;color:#666;">問題を準備しています...</p>';

            Promise.all([gasRun('getTerms', [null]), gasRun('getTermProgress', [currentStudent])])
                .then(function (res) {
                    termAll = res[0] || [];
                    var prog = res[1] || {};
                    termLearnedIds = prog.learned_term_ids || [];

                    var pool = filterTerms(termAll).filter(function (t) {
                        return termLearnedIds.indexOf(t.term_id) !== -1;
                    });

                    if (pool.length < 4) {
                        document.getElementById('termQuizContent').innerHTML =
                            '<p style="text-align:center;color:#c33;margin-bottom:20px;">「覚えた」用語が4語より少ないため、まだクイズを作れません。<br>フラッシュカードで用語を覚えてから挑戦してください。</p>' +
                            '<button class="btn-secondary" onclick="goToTermMenu()">← 用語学習に戻る</button>';
                        return;
                    }

                    shuffleArr(pool);
                    var picked = pool.slice(0, Math.min(10, pool.length));
                    termQuizList = picked.map(function (t) {
                        var others = pool.filter(function (o) { return o.term_id !== t.term_id; });
                        shuffleArr(others);
                        var choices = [t].concat(others.slice(0, 3));
                        shuffleArr(choices);
                        return { term: t, choices: choices };
                    });
                    termQuizIndex = 0;
                    termQuizCorrect = 0;
                    termQuizAnswered = false;
                    renderTermQuiz();
                });
        }

        function renderTermQuiz() {
            if (termQuizIndex >= termQuizList.length) {
                renderTermQuizResult();
                return;
            }
            var item = termQuizList[termQuizIndex];
            var t = item.term;

            var html = '';
            html += '<div class="quiz-top">';
            html += '  <button type="button" class="quiz-abort" onclick="goToTermMenu()">やめて用語学習メニューへ戻る</button>';
            html += '</div>';
            html += '<div class="quiz-progress-bar"><div class="quiz-progress-bar__fill" style="width:' + Math.round((termQuizIndex / termQuizList.length) * 100) + '%"></div></div>';
            html += '<p class="quiz-progress-label">第' + (termQuizIndex + 1) + '問 / 全' + termQuizList.length + '問</p>';
            html += '<div class="quiz-meta">';
            html += '  <span class="quiz-meta__item">ID: ' + escapeHtml(t.term_id || '') + '</span>';
            html += '  <span class="quiz-meta__item">' + escapeHtml(String(t.chapter || '')) + '</span>';
            html += '  <span class="quiz-meta__item quiz-meta__item--select-type">' + escapeHtml(t.level || '') + '</span>';
            html += '</div>';
            html += '<div class="quiz-card">';
            html += '  <p class="quiz-question">「' + escapeHtml(t.term) + '」の説明として正しいものはどれ？</p>';
            html += '  <div id="termQuizChoices">';
            item.choices.forEach(function (c, i) {
                html += '<button type="button" class="term-quiz-choice" onclick="answerTermQuiz(' + i + ')">' + escapeHtml(c.definition || '') + '</button>';
            });
            html += '  </div>';
            html += '  <div id="termQuizJudge"></div>';
            html += '</div>';
            html += '<div id="termQuizNext"></div>';

            document.getElementById('termQuizContent').innerHTML = html;
            termQuizAnswered = false;
        }

        function answerTermQuiz(choiceIndex) {
            if (termQuizAnswered) return;
            termQuizAnswered = true;

            var item = termQuizList[termQuizIndex];
            var t = item.term;
            var chosen = item.choices[choiceIndex];
            var isCorrect = (chosen && chosen.term_id === t.term_id);
            if (isCorrect) termQuizCorrect += 1;

            // 回答ごとに即時記録（章別・難易度別の集計に使う）
            gasRun('postTermQuizResult', [currentStudent, {
                term_id: t.term_id,
                chapter: t.chapter,
                level: t.level,
                judgment: isCorrect ? '正解' : '不正解'
            }]);

            var buttons = document.querySelectorAll('#termQuizChoices .term-quiz-choice');
            buttons.forEach(function (b, i) {
                var c = item.choices[i];
                if (c && c.term_id === t.term_id) {
                    b.style.borderColor = '#28a745';
                    b.style.background = '#eaf7ee';
                } else if (i === choiceIndex) {
                    b.style.borderColor = '#dc3545';
                    b.style.background = '#fdecea';
                }
            });

            document.getElementById('termQuizJudge').innerHTML =
                '<p class="term-judge ' + (isCorrect ? 'term-judge--ok' : 'term-judge--ng') + '">' + (isCorrect ? '正解' : '不正解') + '</p>' +
                '<p class="result-item__explanation-label">正しい説明</p>' +
                '<p class="result-item__explanation">' + escapeHtml(t.definition || '') + '</p>';

            document.getElementById('termQuizNext').innerHTML =
                '<button type="button" class="quiz-next-btn" onclick="nextTermQuiz()">' +
                (termQuizIndex === termQuizList.length - 1 ? '結果を見る' : '次の問題へ') + '</button>';
        }

        function nextTermQuiz() {
            termQuizIndex += 1;
            renderTermQuiz();
            window.scrollTo(0, 0);
        }

        function renderTermQuizResult() {
            var total = termQuizList.length;
            var rate = total > 0 ? Math.round((termQuizCorrect / total) * 1000) / 10 : 0;
            var html = '';
            html += '<div class="result-summary">';
            html += '  <p class="result-summary__label">用語クイズの結果</p>';
            html += '  <p class="result-summary__score">' + termQuizCorrect + ' <span>/ ' + total + '問</span></p>';
            html += '  <p class="result-summary__rate">正答率 ' + rate + '%</p>';
            if (rate < PASS_LINE) {
                html += '  <div class="pass-badge pass-badge--ng">合格ライン（' + PASS_LINE + '%）未達</div>';
            } else {
                html += '  <div class="pass-badge pass-badge--ok">合格ライン（' + PASS_LINE + '%）クリア</div>';
            }
            html += '</div>';
            html += '<button type="button" class="result-back-btn" onclick="goToTermMenu()">← 用語学習に戻る</button>';
            document.getElementById('termQuizContent').innerHTML = html;
            window.scrollTo(0, 0);
        }

        /* ===== 問題インポート ===== */
        var importFileTexts = null;
        function showImport() {
            importFileTexts = null;
            document.getElementById('importResult').innerHTML = '';
            var fi = document.getElementById('importFileInfo'); if (fi) fi.textContent = '';
            var fe = document.getElementById('importFile'); if (fe) fe.value = '';
            showScreen('importScreen');
            window.scrollTo(0, 0);
        }
        function getImportPayload() {
            return (importFileTexts && importFileTexts.length) ? importFileTexts : [];
        }
        function getImportMode() {
            var els = document.querySelectorAll('input[name="importMode"]');
            for (var i = 0; i < els.length; i++) { if (els[i].checked) return els[i].value; }
            return 'append';
        }
        function loadImportFiles(input) {
            var files = input.files;
            if (!files || !files.length) return;
            var info = document.getElementById('importFileInfo');
            info.style.color = '#667eea';
            info.textContent = '読み込み中...';
            var texts = [];
            var remaining = files.length;
            var failed = 0;
            for (var i = 0; i < files.length; i++) {
                (function (idx) {
                    var reader = new FileReader();
                    reader.onload = function (e) {
                        texts[idx] = e.target.result || '';
                        remaining--;
                        if (remaining === 0) done();
                    };
                    reader.onerror = function () {
                        texts[idx] = '';
                        failed++;
                        remaining--;
                        if (remaining === 0) done();
                    };
                    reader.readAsText(files[idx]); // UTF-8
                })(i);
            }
            function done() {
                importFileTexts = texts.filter(function (t) { return t && t.trim(); });
                var names = [];
                for (var k = 0; k < files.length; k++) { names.push(files[k].name); }
                info.style.color = failed ? '#c0392b' : '#28a745';
                info.textContent = '読み込んだファイル：' + names.join('、')
                    + '（' + importFileTexts.length + '個）'
                    + (failed ? ' ※' + failed + '個は失敗' : '')
                    + '／次に「チェック（プレビュー）」を押してください。';
            }
        }
        function renderImportResult(res, isPreview) {
            var box = document.getElementById('importResult');
            if (!res || !res.ok) {
                var msg = (res && res.message) ? res.message : '取り込みに失敗しました';
                box.innerHTML = '<div class="import-result__box import-result__err">⚠ ' + escapeHtml(msg) + '</div>';
                return;
            }
            var errsHtml = '';
            if (res.errors && res.errors.length) {
                errsHtml = '<div style="margin-top:8px;">気になる点:<ul style="margin:6px 0 0 18px;">';
                res.errors.forEach(function (e) { errsHtml += '<li>' + escapeHtml(e) + '</li>'; });
                errsHtml += '</ul></div>';
            }
            var cls = (res.errors && res.errors.length) ? 'import-result__warn' : 'import-result__ok';
            var head = isPreview ? 'プレビュー結果（まだ保存していません）' : '取り込み完了！';
            var body = '問題 ' + res.questionCount + ' 問 ／ 選択肢 ' + res.choiceCount + ' 件';
            if (res.target_exam) { body += '<br>対応試験：' + escapeHtml(res.target_exam); }
            if (!isPreview && res.version) { body += '<br>データ版：v' + res.version + '（' + escapeHtml(res.mode === 'replace' ? '全置換' : '追記') + '）'; }
            box.innerHTML = '<div class="import-result__box ' + cls + '"><b>' + head + '</b><br>' + body + errsHtml + '</div>';
        }
        function previewImport() {
            var arr = getImportPayload();
            if (!arr.length) { alert('ファイルを選ぶか、CSVを貼り付けてください'); return; }
            document.getElementById('importResult').innerHTML = '<div class="loading"></div>';
            gasRun('importAuto', [arr, getImportMode(), true]).then(function (res) {
                renderImportResult(res, true);
            });
        }
        function runImport() {
            var arr = getImportPayload();
            if (!arr.length) { alert('ファイルを選ぶか、CSVを貼り付けてください'); return; }
            var mode = getImportMode();
            if (mode === 'replace' && !confirm('全置換モードです。今ある問題と選択肢を全て消して入れ替えます。よろしいですか？')) { return; }
            document.getElementById('importResult').innerHTML = '<div class="loading"></div>';
            gasRun('importAuto', [arr, mode, false]).then(function (res) {
                renderImportResult(res, false);
            });
        }

        window.addEventListener('DOMContentLoaded', init);

        /* ===== スロースクロール(ゆっくり滑らかなスクロール) ===== */
        (function () {
            var SCROLL_EASE = 0.07; // 小さいほどゆっくり(0.05〜0.15が目安)
            var STEP = 55;          // 1回のホイールで進む量(px)
            var targetY = window.scrollY || window.pageYOffset;
            var animating = false;

            function maxScroll() {
                return Math.max(0, document.documentElement.scrollHeight - window.innerHeight);
            }
            function animate() {
                var current = window.scrollY || window.pageYOffset;
                var diff = targetY - current;
                if (Math.abs(diff) < 0.5) {
                    window.scrollTo(0, targetY);
                    animating = false;
                    return;
                }
                window.scrollTo(0, current + diff * SCROLL_EASE);
                requestAnimationFrame(animate);
            }
            function onWheel(e) {
                if (maxScroll() <= 0) return;
                e.preventDefault();
                var dir = e.deltaY > 0 ? 1 : -1;
                targetY += dir * STEP;
                targetY = Math.max(0, Math.min(targetY, maxScroll()));
                if (!animating) {
                    animating = true;
                    requestAnimationFrame(animate);
                }
            }
            window.addEventListener('wheel', onWheel, { passive: false });
            window.addEventListener('resize', function () {
                targetY = window.scrollY || window.pageYOffset;
            });
        })();
    </script>
</body>
</html>

`)
    .setTitle('模擬試験アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =========================================================
// google.script.run から呼び出す公開関数
// （末尾に _ が付いた関数は google.script.run から呼べないため、
//   ここで公開用のラッパーを用意する）
// =========================================================
function getStudents() {
  return getStudents_();
}
function getQuestions(category, type) {
  return getQuestions_(category, type);
}
function getChoices(questionId) {
  return getChoices_(questionId);
}
function getHistory(studentName) {
  return getHistory_(studentName);
}
function postExamResult(studentName, examData) {
  return postExamResult_(studentName, examData);
}
function getTerms(chapter) {
  return getTerms_(chapter);
}
function getTermProgress(studentName) {
  return getTermProgress_(studentName);
}
function postTermLearned(studentName, termId) {
  return postTermLearned_(studentName, termId);
}
function postTermQuizResult(studentName, quizData) {
  return postTermQuizResult_(studentName, quizData);
}
function importQuestions(questionsCsv, choicesCsv, mode, dryRun) {
  return importQuestions_(questionsCsv, choicesCsv, mode, dryRun);
}
function importKintone(csvText, mode, dryRun) {
  return importKintone_(csvText, mode, dryRun);
}
function importAuto(texts, mode, dryRun) {
  return importAuto_(texts, mode, dryRun);
}
function getDashboardStats(studentName) {
  return getDashboardStats_(studentName);
}
function getWrongQuestionIds(studentName) {
  return getWrongQuestionIds_(studentName);
}
function getTermDashboard(studentName, chapter, level) {
  return getTermDashboard_(studentName, chapter, level);
}
function addStudent(name) {
  return addStudent_(name);
}

// ===== 文字列の正規化（NFKC＋不可視文字除去＋トリム） =====
// スプレッドシートのカテゴリー等に、見た目は同じでも異なる文字コード
// （NFC/NFD、全角半角、ゼロ幅スペース、BOM 等）が混じっていても一致させるため。
function normKey_(s) {
  return String(s == null ? '' : s)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF\u00A0\u3000]/g, '')
    .trim();
}


// ===== 問題ID(question_id)の正規化 =====
// 「1-12」のようなIDは、Googleスプレッドシートに取り込むと日付(2026/1/12)として
// 保存されてしまうことがある。セルによって Date 型だったり日付書式の文字列だったりと
// バラバラになるため、比較・受け渡しの前に必ずこの関数を通して「月-日」表記に戻す。
// 日付ではないID(REAL-013 など)はそのまま(正規化のみ)返す。
function qidKey_(v) {
  if (v instanceof Date) {
    return (v.getMonth() + 1) + '-' + v.getDate();
  }
  var s = normKey_(v);
  var m = s.match(/^(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})(?:[ T].*)?$/);
  if (m) {
    return String(Number(m[2])) + '-' + String(Number(m[3]));
  }
  return s;
}

// ===== 学生マスタ =====
function getStudents_() {
  var sheet = SS.getSheetByName(SHEET_STUDENTS);
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {  // student_name が空でなければ
      result.push({
        student_name: row[0],
        display_order: row[1] || i
      });
    }
  }

  result.sort(function(a, b) { return (a.display_order || 0) - (b.display_order || 0); });
  return result;
}

// 名簿(students)に存在する名前かどうか。
// localStorage を書き換えて名簿外の名前で記録されるのを防ぐために使う。
function rosterHasName_(name) {
  var list = getStudents_();
  for (var i = 0; i < list.length; i++) {
    if (normKey_(list[i].student_name) === normKey_(name)) return true;
  }
  return false;
}

// 受験者を追加する（同姓同名は「田中太郎(営業)」のように区別できる名前にしてもらう方針）
function addStudent_(name) {
  var clean = normKey_(name);
  if (!clean) return { ok: false, message: '氏名を入力してください' };
  if (clean.length > 40) return { ok: false, message: '氏名が長すぎます' };

  var list = getStudents_();
  for (var i = 0; i < list.length; i++) {
    if (normKey_(list[i].student_name) === clean) {
      return { ok: false, message: '同じ名前がすでに登録されています。区別できる言葉を足してください（例：田中太郎(営業)）' };
    }
  }

  var sheet = SS.getSheetByName(SHEET_STUDENTS);
  var maxOrder = 0;
  list.forEach(function (st) {
    var n = Number(st.display_order) || 0;
    if (n > maxOrder) maxOrder = n;
  });
  sheet.appendRow([clean, maxOrder + 1]);
  return { ok: true, student_name: clean };
}

// ===== 問題・選択肢 =====
function getQuestions_(category, type) {
  var sheet = SS.getSheetByName(SHEET_QUESTIONS);
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var qid = row[0];  // question_id

    if (category && normKey_(row[1]) !== normKey_(category)) continue;
    if (type && normKey_(row[2]) !== normKey_(type)) continue;

    result.push({
      question_id: qidKey_(qid),
      question_category: row[1],
      question_type: row[2],
      select_type: row[3],
      question: row[4]
    });
  }

  return result;
}

function getChoices_(questionId) {
  var sheet = SS.getSheetByName(SHEET_CHOICES);
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (qidKey_(row[0]) === qidKey_(questionId)) {
      result.push({
        question_id: qidKey_(row[0]),
        choice_no: row[1],
        choice_text: row[2],
        explanation_text: row[3],
        explanation_url: row[4],
        is_correct: row[5]
      });
    }
  }

  result.sort(function(a, b) { return a.choice_no - b.choice_no; });
  return result;
}

/**
 * 選択肢の一括取得（高速化）
 * 問題IDの配列を受け取り、choicesシートを1回だけ読んで
 * { 問題ID: [選択肢, ...] } の形で返す。
 * 従来は問題1件ごとに getChoices を呼んでいたため、
 * 出題数ぶんの通信とシート全読みが発生して遅かった。
 */
function getChoicesBatch(questionIds) {
  return getChoicesBatch_(questionIds);
}
function getChoicesBatch_(questionIds) {
  var sheet = SS.getSheetByName(SHEET_CHOICES);
  var data = sheet.getDataRange().getValues();

  var want = {};
  (questionIds || []).forEach(function (id) { want[qidKey_(id)] = true; });

  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var key = qidKey_(row[0]);
    if (!want[key]) continue;
    if (!map[key]) map[key] = [];
    map[key].push({
      question_id: key,
      choice_no: row[1],
      choice_text: row[2],
      explanation_text: row[3],
      explanation_url: row[4],
      is_correct: row[5]
    });
  }

  Object.keys(map).forEach(function (k) {
    map[k].sort(function (a, b) { return a.choice_no - b.choice_no; });
  });

  return map;
}

// ===== 受験履歴 =====
function getHistory_(studentName) {
  var sheet = SS.getSheetByName(SHEET_HISTORY);
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[2] === studentName) {  // student_name で検索
      result.push({
        history_id: row[0],
        exam_datetime: row[1],
        student_name: row[2],
        exam_mode: row[3],
        selected_category: row[4],
        selected_type: row[5],
        total_count: row[6],
        total_score: row[7],
        correct_rate: row[8],
        duration_seconds: row[9]
      });
    }
  }

  result.sort(function(a, b) { return new Date(b.exam_datetime) - new Date(a.exam_datetime); });
  return result;
}

function getHistoryResults_(historyId) {
  var sheet = SS.getSheetByName(SHEET_HISTORY_RESULTS);
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] === historyId) {
      result.push({
        history_id: row[0],
        row_no: row[1],
        question_id: qidKey_(row[2]),
        question_text: row[3],
        selected_answer: row[4],
        judgment: row[5],
        explanation_text: row[6],
        explanation_url: row[7]
      });
    }
  }

  result.sort(function(a, b) { return a.row_no - b.row_no; });
  return result;
}

function postExamResult_(studentName, examData) {
  // examData = { exam_mode, category, type, results: [...], duration_seconds }
  if (!rosterHasName_(studentName)) {
    throw new Error('名簿にない受験者名のため記録できません: ' + studentName);
  }

  var historySheet = SS.getSheetByName(SHEET_HISTORY);
  var resultsSheet = SS.getSheetByName(SHEET_HISTORY_RESULTS);

  var now = new Date();
  var historyId = Utilities.formatDate(now, 'JST', 'yyyyMMdd_HHmmss') + '_' + Math.floor(Math.random() * 10000);

  var correctCount = 0;
  examData.results.forEach(function(r) {
    if (r.judgment === '正解') correctCount++;
  });
  var correctRate = Math.round((correctCount / examData.results.length) * 100);

  historySheet.appendRow([
    historyId,
    now,
    studentName,
    examData.exam_mode,
    examData.category || '',
    examData.type || '',
    examData.results.length,
    correctCount,
    correctRate,
    examData.duration_seconds || 0
  ]);

  examData.results.forEach(function(r, i) {
    resultsSheet.appendRow([
      historyId,
      i + 1,
      r.question_id,
      r.question_text,
      r.selected_answer,
      r.judgment,
      r.explanation_text || '',
      r.explanation_url || ''
    ]);
  });

  return true;
}

// ===== 用語マスタ =====
function getTerms_(chapter) {
  var sheet = SS.getSheetByName(SHEET_TERMS);
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (chapter && String(row[1]).split('章')[0] !== String(chapter)) continue;

    result.push({
      term_id: row[0],
      chapter: row[1],
      term: row[2],
      definition: row[3],
      level: row[4],
      explanation_url: row[5]
    });
  }

  result.sort(function(a, b) { return a.chapter - b.chapter; });
  return result;
}

// ===== 学習進捗 =====
function getTermProgress_(studentName) {
  var termsSheet = SS.getSheetByName(SHEET_TERMS);
  var progressSheet = SS.getSheetByName(SHEET_TERM_PROGRESS);

  var allTerms = termsSheet.getDataRange().getValues();
  var totalTerms = allTerms.length - 1;

  var progressData = progressSheet.getDataRange().getValues();
  var learnedTerms = [];

  for (var i = 1; i < progressData.length; i++) {
    var row = progressData[i];
    if (row[1] === studentName) {  // student_name で検索
      learnedTerms.push(row[0]);  // term_id
    }
  }

  var learnedCount = learnedTerms.length;
  var percentage = totalTerms > 0 ? Math.round((learnedCount / totalTerms) * 100) : 0;

  return {
    total_terms: totalTerms,
    learned_count: learnedCount,
    percentage: percentage,
    learned_term_ids: learnedTerms
  };
}

function postTermLearned_(studentName, termId) {
  if (!rosterHasName_(studentName)) {
    throw new Error('名簿にない受験者名のため記録できません: ' + studentName);
  }
  var progressSheet = SS.getSheetByName(SHEET_TERM_PROGRESS);
  var data = progressSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === termId && data[i][1] === studentName) {
      return false;  // 既に記録済み
    }
  }

  var now = new Date();
  progressSheet.appendRow([
    termId,
    studentName,
    now
  ]);

  return true;
}

// ===== 用語クイズ履歴 =====
function postTermQuizResult_(studentName, answerData) {
  // answerData = { term_id, chapter, level, judgment }
  // 章別・難易度別の正答率を出せるよう「1回答=1レコード」で記録する
  if (!rosterHasName_(studentName)) {
    throw new Error('名簿にない受験者名のため記録できません: ' + studentName);
  }

  var sheet = SS.getSheetByName(SHEET_TERM_QUIZ_HISTORY);
  var now = new Date();
  var quizId = Utilities.formatDate(now, 'JST', 'yyyyMMdd_HHmmss') + '_' + Math.floor(Math.random() * 10000);

  sheet.appendRow([
    quizId,
    now,
    studentName,
    answerData.term_id,
    answerData.chapter,
    answerData.level,
    answerData.judgment
  ]);

  return true;
}

// =========================================================
// ダッシュボード集計（キントーン版 menu.js と同じ考え方）
// =========================================================

// 本人の受験履歴と解答明細から、カテゴリー別・出題タイプ別の正答率を集計する。
// 集計は「全回答ベース」（受験回ごとの平均ではなく、1問1問を数える）。
function getDashboardStats_(studentName) {
  var hData = SS.getSheetByName(SHEET_HISTORY).getDataRange().getValues();
  var rData = SS.getSheetByName(SHEET_HISTORY_RESULTS).getDataRange().getValues();
  var qData = SS.getSheetByName(SHEET_QUESTIONS).getDataRange().getValues();

  var myHistoryIds = {};
  var examCount = 0, totalCount = 0, totalScore = 0;
  for (var i = 1; i < hData.length; i++) {
    if (normKey_(hData[i][2]) !== normKey_(studentName)) continue;
    myHistoryIds[String(hData[i][0])] = true;
    examCount++;
    totalCount += Number(hData[i][6]) || 0;
    totalScore += Number(hData[i][7]) || 0;
  }

  // question_id → カテゴリー / 出題タイプ
  var qMap = {};
  for (var j = 1; j < qData.length; j++) {
    qMap[qidKey_(qData[j][0])] = { category: normKey_(qData[j][1]), type: normKey_(qData[j][2]) };
  }

  var CATS = ['アプリ', 'スペース', '管理者', 'その他'];
  var TYPES = ['知識問題', '実務問題', 'アプリストア問題'];
  var catAgg = {}, typeAgg = {};
  CATS.forEach(function (c) { catAgg[c] = { total: 0, done: 0 }; });
  TYPES.forEach(function (t) { typeAgg[t] = { total: 0, done: 0 }; });

  for (var k = 1; k < rData.length; k++) {
    var row = rData[k];
    if (!myHistoryIds[String(row[0])]) continue;
    var q = qMap[qidKey_(row[2])];
    if (!q) continue;
    var ok = (normKey_(row[5]) === '正解');
    if (catAgg[q.category]) {
      catAgg[q.category].total++;
      if (ok) catAgg[q.category].done++;
    }
    if (typeAgg[q.type]) {
      typeAgg[q.type].total++;
      if (ok) typeAgg[q.type].done++;
    }
  }

  return {
    exam_count: examCount,
    total_count: totalCount,
    total_score: totalScore,
    avg_rate: totalCount > 0 ? Math.round((totalScore / totalCount) * 100) : 0,
    by_category: CATS.map(function (c) { return { key: c, total: catAgg[c].total, done: catAgg[c].done }; }),
    by_type: TYPES.map(function (t) { return { key: t, total: typeAgg[t].total, done: typeAgg[t].done }; })
  };
}

// 復習モード用：本人が過去に一度でも不正解だった問題IDの一覧
// （その後に正解していても対象に含める方針）
function getWrongQuestionIds_(studentName) {
  var hData = SS.getSheetByName(SHEET_HISTORY).getDataRange().getValues();
  var rData = SS.getSheetByName(SHEET_HISTORY_RESULTS).getDataRange().getValues();

  var myHistoryIds = {};
  for (var i = 1; i < hData.length; i++) {
    if (normKey_(hData[i][2]) === normKey_(studentName)) myHistoryIds[String(hData[i][0])] = true;
  }

  var seen = {};
  var ids = [];
  for (var k = 1; k < rData.length; k++) {
    var row = rData[k];
    if (!myHistoryIds[String(row[0])]) continue;
    if (normKey_(row[5]) !== '不正解') continue;
    // Date オブジェクトのまま返すと google.script.run の戻り値変換に失敗するため、
    // 必ず正規化した文字列を返すこと。
    var id = qidKey_(row[2]);
    if (!seen[id]) {
      seen[id] = true;
      ids.push(id);
    }
  }
  return ids;
}

// 用語学習のダッシュボード：章別・難易度別の達成率＋用語クイズ正答率
function getTermDashboard_(studentName, chapter, level) {
  var termData = SS.getSheetByName(SHEET_TERMS).getDataRange().getValues();
  var progData = SS.getSheetByName(SHEET_TERM_PROGRESS).getDataRange().getValues();
  var quizData = SS.getSheetByName(SHEET_TERM_QUIZ_HISTORY).getDataRange().getValues();

  // 覚えた用語ID（本人分）
  var learnedSet = {};
  for (var p = 1; p < progData.length; p++) {
    if (normKey_(progData[p][1]) === normKey_(studentName)) learnedSet[normKey_(progData[p][0])] = true;
  }

  var chapAgg = {}, levelAgg = {};
  var LEVELS = ['基礎', '応用'];
  LEVELS.forEach(function (lv) { levelAgg[lv] = { total: 0, done: 0 }; });

  var filteredTotal = 0, filteredLearned = 0;

  for (var i = 1; i < termData.length; i++) {
    var row = termData[i];
    if (!row[0]) continue;
    var ch = String(row[1] || '');
    var lv = normKey_(row[4]);
    var learned = !!learnedSet[normKey_(row[0])];

    if (!chapAgg[ch]) chapAgg[ch] = { total: 0, done: 0 };
    chapAgg[ch].total++;
    if (learned) chapAgg[ch].done++;

    if (levelAgg[lv]) {
      levelAgg[lv].total++;
      if (learned) levelAgg[lv].done++;
    }

    // 絞り込み条件に合うものだけを上部カードの達成率に使う
    if (chapter && ch.split('章')[0] !== String(chapter)) continue;
    if (level && lv !== normKey_(level)) continue;
    filteredTotal++;
    if (learned) filteredLearned++;
  }

  // 用語クイズ（1回答=1レコード）
  var quizChap = {}, quizLevel = {};
  LEVELS.forEach(function (lv) { quizLevel[lv] = { total: 0, done: 0 }; });
  for (var q = 1; q < quizData.length; q++) {
    var qrow = quizData[q];
    if (normKey_(qrow[2]) !== normKey_(studentName)) continue;
    var qch = String(qrow[4] || '');
    var qlv = normKey_(qrow[5]);
    var ok = (normKey_(qrow[6]) === '正解');
    if (!quizChap[qch]) quizChap[qch] = { total: 0, done: 0 };
    quizChap[qch].total++;
    if (ok) quizChap[qch].done++;
    if (quizLevel[qlv]) {
      quizLevel[qlv].total++;
      if (ok) quizLevel[qlv].done++;
    }
  }

  function toRows_(agg, suffix) {
    return Object.keys(agg).sort().map(function (k) {
      return { key: k + (suffix || ''), total: agg[k].total, done: agg[k].done };
    });
  }

  return {
    total: filteredTotal,
    learned: filteredLearned,
    rate: filteredTotal > 0 ? Math.round((filteredLearned / filteredTotal) * 100) : 0,
    by_chapter: toRows_(chapAgg, ''),
    by_level: LEVELS.map(function (lv) { return { key: lv, total: levelAgg[lv].total, done: levelAgg[lv].done }; }),
    quiz_by_chapter: toRows_(quizChap, ''),
    quiz_by_level: LEVELS.map(function (lv) { return { key: lv, total: quizLevel[lv].total, done: quizLevel[lv].done }; })
  };
}

// =========================================================
// 問題インポート（CSV → questions / choices）
// =========================================================
var SHEET_DATA_VERSION = 'data_version';

function importQuestions_(questionsCsv, choicesCsv, mode, dryRun) {
  try {
    mode = (mode === 'replace') ? 'replace' : 'append';
    var errors = [];

    var qRows = parseImportCsv_(questionsCsv, 5);
    var cRows = parseImportCsv_(choicesCsv, 6);

    // 問題：question_id をキーに整形
    var qMap = {};
    var qOrder = [];
    qRows.forEach(function (r) {
      var id = String(r[0]).trim();
      if (!id) { errors.push('問題: question_id が空の行があります'); return; }
      if (!qMap[id]) qOrder.push(id);
      qMap[id] = [id, r[1], r[2], r[3], r[4]];
    });

    // 選択肢：question_id ごとにまとめる
    var cByQ = {};
    cRows.forEach(function (r) {
      var id = String(r[0]).trim();
      if (!id) { errors.push('選択肢: question_id が空の行があります'); return; }
      if (!cByQ[id]) cByQ[id] = [];
      var no = Number(r[1]) || (cByQ[id].length + 1);
      var ok = normCorrect_(r[5]);
      cByQ[id].push([id, no, r[2], r[3], r[4], ok ? '○' : '']);
    });

    // 相互チェック（警告のみ・取り込みは止めない）
    qOrder.forEach(function (id) {
      var cs = cByQ[id] || [];
      if (cs.length === 0) errors.push('問題 ' + id + ' に選択肢がありません');
      else if (!cs.some(function (c) { return c[5] === '○'; })) errors.push('問題 ' + id + ' に正解(○)がありません');
    });
    Object.keys(cByQ).forEach(function (id) {
      if (!qMap[id]) errors.push('選択肢の question_id ' + id + ' に対応する問題がありません');
    });

    var questionCount = qOrder.length;
    var choiceCount = 0;
    qOrder.forEach(function (id) { choiceCount += (cByQ[id] || []).length; });

    if (dryRun) {
      return { ok: true, questionCount: questionCount, choiceCount: choiceCount, errors: errors, mode: mode };
    }

    var qSheet = SS.getSheetByName(SHEET_QUESTIONS);
    var cSheet = SS.getSheetByName(SHEET_CHOICES);

    if (mode === 'replace') {
      clearDataRows_(qSheet);
      clearDataRows_(cSheet);
    } else {
      deleteRowsByKey_(qSheet, 0, qOrder);
      deleteRowsByKey_(cSheet, 0, Object.keys(cByQ));
    }

    var qOut = qOrder.map(function (id) { return qMap[id]; });
    if (qOut.length) {
      qSheet.getRange(qSheet.getLastRow() + 1, 1, qOut.length, 5).setValues(qOut);
    }

    var cOut = [];
    qOrder.forEach(function (id) { (cByQ[id] || []).forEach(function (c) { cOut.push(c); }); });
    Object.keys(cByQ).forEach(function (id) {
      if (qOrder.indexOf(id) === -1) { cByQ[id].forEach(function (c) { cOut.push(c); }); }
    });
    if (cOut.length) {
      cSheet.getRange(cSheet.getLastRow() + 1, 1, cOut.length, 6).setValues(cOut);
    }

    var version = bumpDataVersion_(questionCount);

    return { ok: true, questionCount: questionCount, choiceCount: cOut.length, errors: errors, mode: mode, version: version };
  } catch (err) {
    return { ok: false, message: (err && err.message) ? err.message : String(err) };
  }
}

function parseImportCsv_(csv, minCols) {
  if (!csv || !csv.trim()) return [];
  var data = Utilities.parseCsv(csv);
  if (!data || data.length === 0) return [];
  var start = 0;
  var first = (data[0][0] || '').toString().trim().toLowerCase();
  if (first === 'question_id') start = 1; // 見出し行はスキップ
  var rows = [];
  for (var i = start; i < data.length; i++) {
    var row = data[i];
    if (!row || row.join('').trim() === '') continue; // 空行スキップ
    while (row.length < minCols) row.push('');
    rows.push(row);
  }
  return rows;
}

function normCorrect_(v) {
  if (v === true) return true;
  if (typeof v === 'number') return v === 1;
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return s === '○' || s === '◯' || s === '1' || s === 'true' || s === '✓' || s === 'yes' || s === '正解' || s === 'o';
}

function clearDataRows_(sheet) {
  var last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
}

function deleteRowsByKey_(sheet, keyColIndex, keys) {
  if (!keys || keys.length === 0) return;
  var keySet = {};
  keys.forEach(function (k) { keySet[String(k).trim()] = true; });
  var last = sheet.getLastRow();
  if (last < 2) return;
  var vals = sheet.getRange(2, 1, last - 1, keyColIndex + 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    var key = String(vals[i][keyColIndex]).trim();
    if (keySet[key]) sheet.deleteRow(i + 2);
  }
}

// 知識問題データの最新バージョンを返す（メニュー下部の表示用）
function getDataVersion() {
  return getDataVersion_();
}

function getDataVersion_() {
  var sheet = SS.getSheetByName(SHEET_DATA_VERSION);
  if (!sheet) return null;
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var row = sheet.getRange(last, 1, 1, 4).getValues()[0];
  var at = row[2];
  var atText = "";
  if (at instanceof Date) {
    atText = Utilities.formatDate(at, "JST", "yyyy/MM/dd");
  } else if (at) {
    atText = String(at);
  }
  return { version: row[0], target_exam: row[1], imported_at: atText, imported_count: row[3] };
}

function bumpDataVersion_(count, targetExam) {
  var sheet = SS.getSheetByName(SHEET_DATA_VERSION);
  if (!sheet) {
    sheet = SS.insertSheet(SHEET_DATA_VERSION);
    sheet.appendRow(['version', 'target_exam', 'imported_at', 'imported_count']);
  }
  var last = sheet.getLastRow();
  var version = 1;
  if (last > 1) {
    var prev = sheet.getRange(last, 1).getValue();
    version = (Number(prev) || 0) + 1;
  }
  sheet.appendRow([version, (targetExam || 'kintone'), new Date(), count]);
  return version;
}

// =========================================================
// kintone 練習問題CSV → questions / choices へ変換して取込
//   1行＝1問（選択肢A〜Dが横並び、「正答」の記号で正解判定）
// =========================================================
function importKintone_(csvText, mode, dryRun) {
  try {
    mode = (mode === 'replace') ? 'replace' : 'append';
    var errors = [];
    if (!csvText || !csvText.trim()) return { ok: false, message: 'CSVが空です' };

    var data = Utilities.parseCsv(csvText);
    var qMap = {}, qOrder = [], cByQ = {};
    var LETTERS = ['A', 'B', 'C', 'D'];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row) continue;
      var c0 = String(row[0] == null ? '' : row[0]).trim();
      if (c0 === '') continue;                 // 空行
      if (c0 === '練習問題セット') continue;    // 見出し行（複数ファイル対応）
      if (row.length < 12) { errors.push((i + 1) + '行目: 列が足りません（スキップ）'); continue; }

      var setNo = c0;
      var qno = String(row[1] == null ? '' : row[1]).trim();
      var qid = setNo + '-' + qno;
      var category = normCategory_(row[3]);
      var correct = String(row[5] == null ? '' : row[5]).trim().toUpperCase();
      var question = row[7];

      if (!qMap[qid]) qOrder.push(qid);
      qMap[qid] = [qid, category, '', (correct.length > 1 ? '複数選択' : '単一選択'), question];

      cByQ[qid] = [];
      for (var j = 0; j < 4; j++) {
        var text = row[8 + j];
        if (text === undefined || String(text).trim() === '') continue;  // 選択肢が無ければ飛ばす
        var helpTitle = row[12 + j * 2];
        var helpUrl = row[13 + j * 2];
        var textRef = row[20 + j];
        var expl = '';
        if (textRef && String(textRef).trim()) expl += '【' + String(textRef).trim() + '】';
        if (helpTitle && String(helpTitle).trim()) expl += String(helpTitle).trim();
        var isC = (correct.indexOf(LETTERS[j]) !== -1) ? '○' : '';
        cByQ[qid].push([qid, j + 1, text, expl, helpUrl || '', isC]);
      }
      if (!cByQ[qid].some(function (c) { return c[5] === '○'; })) {
        errors.push('問題 ' + qid + ' に正解(○)がありません（正答欄=「' + correct + '」）');
      }
    }

    return writeImport_(qMap, qOrder, cByQ, mode, dryRun, errors);
  } catch (err) {
    return { ok: false, message: (err && err.message) ? err.message : String(err) };
  }
}

function normCategory_(v) {
  var s = String(v == null ? '' : v).trim();
  if (s === '管理者設定') return '管理者';   // アプリのカテゴリー名に合わせる
  return s;
}

// =========================================================
// 複数ファイルを渡すと、各ファイルの見出しから形式を自動判別して取込
//   - 問題マスタ(native questions): question_id,question_category,question_type,select_type,question
//   - 選択肢(native choices):       question_id,choice_no,choice_text,explanation_text,explanation_url,is_correct
//   - 練習問題セット(kintone):       練習問題セット,設問,…,選択肢A〜D,…（正答の記号で判定）
// =========================================================
function importAuto_(texts, mode, dryRun) {
  try {
    mode = (mode === 'replace') ? 'replace' : 'append';
    if (typeof texts === 'string') texts = [texts];
    if (!texts || !texts.length) return { ok: false, message: 'ファイル/CSVが空です' };

    var errors = [];
    var qMap = {}, order = [], cByQ = {};
    var LETTERS = ['A', 'B', 'C', 'D'];
    var targetExam = '';   // 練習問題CSVの「対応試験」列（最初に見つかった値を使う）

    texts.forEach(function (text, fi) {
      if (!text || !String(text).trim()) return;
      var data = Utilities.parseCsv(text);
      if (!data.length) return;
      var hi = 0;
      while (hi < data.length && (!data[hi] || data[hi].join('').trim() === '')) hi++;
      if (hi >= data.length) return;
      var header = data[hi].map(function (x) { return String(x == null ? '' : x).trim(); });
      var h0 = header[0];
      var kind;
      if (h0 === '練習問題セット') kind = 'practice';
      else if (h0 === 'question_id') kind = (header.indexOf('choice_text') !== -1 || header.indexOf('choice_no') !== -1) ? 'choices' : 'questions';
      else { errors.push('ファイル' + (fi + 1) + '：形式を判別できません（先頭列=「' + h0 + '」）'); return; }

      for (var i = hi + 1; i < data.length; i++) {
        var row = data[i];
        if (!row) continue;
        var first = String(row[0] == null ? '' : row[0]).trim();
        if (first === '') continue;

        if (kind === 'practice') {
          if (first === '練習問題セット') continue;
          if (row.length < 12) continue;
          if (!targetExam) {
            var teIdx = header.indexOf('対応試験');
            if (teIdx >= 0 && row[teIdx] != null && String(row[teIdx]).trim() !== '') targetExam = String(row[teIdx]).trim();
          }
          var qid = first + '-' + String(row[1] == null ? '' : row[1]).trim();
          var correct = String(row[5] == null ? '' : row[5]).trim().toUpperCase();
          if (!qMap[qid]) order.push(qid);
          qMap[qid] = [qid, normCategory_(row[3]), '', (correct.length > 1 ? '複数選択' : '単一選択'), row[7]];
          cByQ[qid] = [];
          for (var j = 0; j < 4; j++) {
            var ct = row[8 + j];
            if (ct === undefined || String(ct).trim() === '') continue;
            var expl = '';
            var tr = row[20 + j], htl = row[12 + j * 2];
            if (tr && String(tr).trim()) expl += '【' + String(tr).trim() + '】';
            if (htl && String(htl).trim()) expl += String(htl).trim();
            cByQ[qid].push([qid, j + 1, ct, expl, row[13 + j * 2] || '', (correct.indexOf(LETTERS[j]) !== -1 ? '○' : '')]);
          }
        } else if (kind === 'questions') {
          var id = first;
          if (!qMap[id]) order.push(id);
          qMap[id] = [id, normCategory_(row[1]), (row[2] == null ? '' : row[2]), (row[3] == null ? '' : row[3]), (row[4] == null ? '' : row[4])];
        } else if (kind === 'choices') {
          var cid = first;
          if (!cByQ[cid]) cByQ[cid] = [];
          var no = Number(row[1]) || (cByQ[cid].length + 1);
          cByQ[cid].push([cid, no, (row[2] == null ? '' : row[2]), (row[3] == null ? '' : row[3]), (row[4] == null ? '' : row[4]), (normCorrect_(row[5]) ? '○' : '')]);
        }
      }
    });

    var seen = {}, qOrder = [];
    order.forEach(function (id) { if (!seen[id]) { seen[id] = 1; qOrder.push(id); } });

    // 選択肢が1件も無い問題は、まとめて分かりやすい1行にする
    var noChoice = [], noCorrect = [];
    qOrder.forEach(function (id) {
      var cs = cByQ[id] || [];
      if (cs.length === 0) noChoice.push(id);
      else if (!cs.some(function (c) { return c[5] === '○'; })) noCorrect.push(id);
    });
    var counts = { noChoice: noChoice.length, noCorrect: noCorrect.length };

    return writeImport_(qMap, qOrder, cByQ, mode, dryRun, errors, targetExam, counts);
  } catch (err) {
    return { ok: false, message: (err && err.message) ? err.message : String(err) };
  }
}

// qMap/qOrder/cByQ を questions/choices シートへ書き込む共通処理
function writeImport_(qMap, qOrder, cByQ, mode, dryRun, errors, targetExam, counts) {
  errors = errors || [];
  targetExam = targetExam || '';
  counts = counts || {};

  // 問題と選択肢の食い違いは、管理者にできることが「取り直して選び直す」だけなので
  // IDを並べず、やることを1行で伝えて状況だけ括弧で添える
  var orphan = 0;
  Object.keys(cByQ).forEach(function (id) { if (!qMap[id]) orphan++; });
  var detail = [];
  if (orphan) detail.push('対応する問題が無い選択肢 ' + orphan + ' 件分');
  if (counts.noChoice) detail.push('選択肢が無い問題 ' + counts.noChoice + ' 件');
  if (counts.noCorrect) detail.push('正解(○)が無い問題 ' + counts.noCorrect + ' 件');
  if (detail.length) {
    errors.push('CSVの内容がそろっていません。ダウンロードし直して、すべてのファイルを選び直してください（' + detail.join(' ／ ') + '）');
  }

  var questionCount = qOrder.length;
  var choiceCount = 0;
  qOrder.forEach(function (id) { choiceCount += (cByQ[id] || []).length; });

  if (dryRun) {
    return { ok: true, questionCount: questionCount, choiceCount: choiceCount, errors: errors, mode: mode, target_exam: targetExam };
  }

  var qSheet = SS.getSheetByName(SHEET_QUESTIONS);
  var cSheet = SS.getSheetByName(SHEET_CHOICES);
  if (!qSheet) return { ok: false, message: 'questions シートが見つかりません' };
  if (!cSheet) return { ok: false, message: 'choices シートが見つかりません' };

  if (mode === 'replace') {
    clearDataRows_(qSheet);
    clearDataRows_(cSheet);
  } else {
    deleteRowsByKey_(qSheet, 0, qOrder);
    deleteRowsByKey_(cSheet, 0, Object.keys(cByQ));
  }

  var qOut = qOrder.map(function (id) { return qMap[id]; });
  if (qOut.length) {
    var qStart = qSheet.getLastRow() + 1;
    qSheet.getRange(qStart, 1, qOut.length, 1).setNumberFormat('@'); // question_id列を文字列扱いにして「3-1→日付」化を防ぐ
    qSheet.getRange(qStart, 1, qOut.length, 5).setValues(qOut);
  }

  var cOut = [];
  qOrder.forEach(function (id) { (cByQ[id] || []).forEach(function (c) { cOut.push(c); }); });
  Object.keys(cByQ).forEach(function (id) {
    if (qOrder.indexOf(id) === -1) { cByQ[id].forEach(function (c) { cOut.push(c); }); }
  });
  if (cOut.length) {
    var cStart = cSheet.getLastRow() + 1;
    cSheet.getRange(cStart, 1, cOut.length, 1).setNumberFormat('@'); // question_id列を文字列扱いに
    cSheet.getRange(cStart, 1, cOut.length, 6).setValues(cOut);
  }

  var version = bumpDataVersion_(questionCount, targetExam);
  return { ok: true, questionCount: questionCount, choiceCount: cOut.length, errors: errors, mode: mode, version: version, target_exam: targetExam };
}
