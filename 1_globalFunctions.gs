const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName('list');
const lastRow = sheet.getRange('B5').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
const triggerMode = sheet.getRange('I2').getValue();
let loadingAnimation = HtmlService.createTemplateFromFile('loadingAnimation');
const closeModalDialog = HtmlService.createHtmlOutput('<script>google.script.host.close()</script>');

function setup(e) { //初期設定
  const ui = SpreadsheetApp.getUi();
  const confirmation1 = ui.alert('トリガーを設定します','あなたはこのシートで管理するフォームの編集権限を持っていますか？',ui.ButtonSet.YES_NO_CANCEL)

  if (confirmation1 === ui.Button.NO) {
    ui.alert('トリガーを設定できません','このシートで管理するフォームの編集権限を持っている人に、トリガーの設定を依頼してください。',ui.ButtonSet.OK);
    return;
  } else if (confirmation1 === ui.Button.YES) {
    const confirmation2 = ui.alert('トリガーを設定します',
      'これからプログラムに付与する権限により、以下の操作が実行されます。なお、他の人がこれらの操作を実行した場合でも、ログにはあなたが実行したものとして記録される場合があります。よろしいですか？\n\n' + 
      '● モード(自動更新/一時停止)の切り替え\n' +
      '● リストの内容が更新された際のフォームへの反映\n' + 
      '● フォーム回答の受付開始/終了の自動更新',
    ui.ButtonSet.YES_NO);

    if (confirmation2 === ui.Button.YES) {
      loadingAnimation.message = 'トリガーを設定しています。';
      ui.showModalDialog(loadingAnimation.evaluate().setWidth(400).setHeight(320), "処理中");
      setTriggers();

      const drawings = sheet.getDrawings()
      for (const drawing of drawings) {
        if (drawing.getOnAction() == 'setup') {
          drawing.remove();
          break;
        }
      }

      ui.showModalDialog(closeModalDialog, '処理が終了しました');
    } else {
      ss.toast('処理を中断しました');
    }
  }
}

function setTriggers() { //このシートのトリガーを設定
  deleteTriggers();
  ScriptApp.newTrigger('updateFormsFromSheet').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('changeTriggerMode').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('startOrStopResponse').timeBased().everyMinutes(1).create();
  ss.toast('トリガーを設定しました');
}

function deleteTriggers(){
  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}

function searchRangeOfValue(value, range) {
  if (!value || !range) return;
  console.log('Search for: ' + value);

  const ranges = range.createTextFinder(value).matchEntireCell(true).findAll().map(range => range);
  for (let i = 0; i < ranges.length; i++) {
    console.log('Range[' + i + ']: ' + ranges[i].getA1Notation());
  }

  if (ranges.length < 1) {
    return null
  } else {
    return ranges;
  }
}

function extractFileId(url) {
  console.log('Url: ' + url);

  try {
    if (/^[-\w]{25,}$/.test(url)) {
        return url; // ファイルIDだけが渡されるケース
      }

    const patterns = [
      /\/d\/([-\w]{25,})/, // "/d/" パターン
      /id=([-\w]{25,})/,   // "id=" パターン
      /\/open\?id=([-\w]{25,})/, // "/open?id=" パターン
      /\/file\/d\/([-\w]{25,})/, // "/file/d/" パターン
      /drive.google.com\/uc\?export=download&id=([-\w]{25,})/, // ダウンロードリンク
      /\/folders\/([-\w]{25,})/, // フォルダの場合のパターン
      /spreadsheets\/d\/([-\w]{25,})/, // Googleスプレッドシートの場合
      /document\/d\/([-\w]{25,})/, // Googleドキュメントの場合
      /presentation\/d\/([-\w]{25,})/, // Googleスライドの場合
    ];

    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) return match[1]; // マッチした場合、IDを返す
    }
  } catch {
    return null;
  }
}