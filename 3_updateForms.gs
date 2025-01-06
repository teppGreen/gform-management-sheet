function updateFormsFromSheet(e) { //編集時実行
  try {
    const ui = SpreadsheetApp.getUi();
    const editedRange = e.range;
    const editedColumn = editedRange.getColumn();
    const editedRow =  editedRange.getRow();
    const sheetName = e.source.getActiveSheet().getName();

    if (editedRow < 6 || sheetName !== 'list') return;
    
    const formUrl = sheet.getRange(editedRow,5).getValue();
    const formId = extractFileId(formUrl);
    const form = FormApp.openById(formId);

    const formFile = DriveApp.getFileById(formId);
    const formName = formFile.getName();
    const isOverLength = formName.length > 24 ? '…' : '';
    const formName_short = formName.substr(1,24) + isOverLength;

    const startResponseDate_range = sheet.getRange(editedRow,7);
      let startResponseDate = startResponseDate_range.getValue();
    const stopResponseDate_range = sheet.getRange(editedRow,8);
      let stopResponseDate = stopResponseDate_range.getValue();
    const now = new Date();

    const isAcceptingResponses_range = sheet.getRange(editedRow, 2);
    const isAcceptingResponses = isAcceptingResponses_range.getValue();
    let switchTo, switchFrom, beforeOrAfter;

    if (triggerMode === '一時停止') return;

    if (editedColumn === 2) { //回答受付中
      console.log('e.value:' + e.value);
      if (isAcceptingResponses) {
        switchTo = '開始';
        if (startResponseDate && now < startResponseDate) { 
          switchFrom = '開始';
          beforeOrAfter = '前';
        } else if (stopResponseDate && now > stopResponseDate) {
          switchFrom = '終了';
          beforeOrAfter = '後';
        }
      } else {
        switchTo = '終了';
        if (stopResponseDate && now < stopResponseDate) { 
          switchFrom = '終了';
          beforeOrAfter = '前';
        }
      }

      if (beforeOrAfter) {
        const confirmation = ui.alert('入力が矛盾しています', '現在日時がシートに記入されている回答受付' + switchFrom + '日時よりも' + beforeOrAfter + 'ですが、今すぐ回答受付を' + switchTo + 'しますか？', ui.ButtonSet.YES_NO);
        if (confirmation !== ui.Button.YES) {
          editedRange.setValue(e.oldValue);
          ss.toast('処理を中断しました');
          return;
        }
      }
      
      form.setAcceptingResponses(isAcceptingResponses);
      if (switchTo === '開始') {
        startResponseDate_range.setValue(now);
        if (switchFrom === '終了') {
          stopResponseDate_range.clearContent();
        }
      } else if (switchTo === '終了') {
        stopResponseDate_range.setValue(now);
      }
      ss.toast(formName, '回答受付を' + switchTo + 'しました');
    }

    if (editedColumn === 3) {//ファイル名
      formFile.setName(e.value);
      ss.toast('元のファイル名: ' + formName, 'ファイル名を変更しました');
    }

    if (editedColumn === 4) { //タイトル
      form.setTitle(e.value);
      ss.toast('元のタイトル: ' + formName, 'タイトルを変更しました');
    }

    if (e.value >= now) { //回答受付開始・終了日時
      if (editedColumn === 7 && !isAcceptingResponses) {
        switchTo = '開始';
      } else if (editedColumn === 8 && isAcceptingResponses) {
        switchTo = '終了';
      }

      const confirmation = ui.alert('過去の日時が指定されました', '今すぐ回答受付を' + switchTo + 'しますか？', ui.ButtonSet.YES_NO); 
      if (confirmation !== ui.Button.YES) {
        editedRange.setValue(e.oldValue);
        ss.toast('処理を中断しました');
        return;
      } else {
        form.setAcceptingResponses(!isAcceptingResponses);
        isAcceptingResponses_range.setValue(!isAcceptingResponses);
        ss.toast(formName, '回答受付を' + switchTo + 'しました');
      }
    }
  } catch(error) {
    console.log('Error: ' + error.message);
    ss.toast(error.message, 'エラー');
  }
}

function changeTriggerMode(e) {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheetName = e.source.getActiveSheet().getName();
    const editedRange = e.range;
    const firstRow = 6;
    const rowLength = lastRow - 5;

    if (sheetName !== 'list' || editedRange.getA1Notation() != 'I2' || rowLength < 1) return;

    loadingAnimation.message = 'モードを変更しています。';
    ui.showModalDialog(loadingAnimation.evaluate().setWidth(400).setHeight(320), "処理中");

    if (e.value === '一時停止') {
      const confirmation = ui.alert('モードが「一時停止」に変更されました', 
        '● 一時停止中はリストの更新内容がフォームに反映されません\n' + 
        '● 回答受付状況(回答受付中のチェックボックス)と回答受付開始/終了日時の矛盾は修正されません\n' + 
        '● 回答の受付開始/終了は自動で行われません\n' + 
        '● 再び「自動更新」にすると、一時停止中の更新内容が全て反映されます。\n\n' + 
        'よろしいですか？',
        ui.ButtonSet.YES_NO);
      if (confirmation === ui.Button.NO) {
        editedRange.setValue(e.oldValue);
      }
    }

    const datas = sheet.getRange(firstRow, 2, rowLength, 8).getValues();
    let forms = new Array();
    let formFiles = new Array();
    let flag;

    for (const data of datas) {
      const formUrl = data[3];
      const formId = extractFileId(formUrl);
      forms.push(FormApp.openById(formId));
      formFiles.push(DriveApp.getFileById(formId));
    }

    if (e.value === '自動反映') {
      const confirmation = ui.alert('モードが「自動反映」に変更されました', 
        '一時停止中のリストの更新内容を今すぐ反映しますか？\n\n' + 
        '「はい/YES」を押した場合は、リストの回答受付状態(回答受付中のチェックボックス)/ファイル名/タイトルがフォームに反映されます。\n' + 
        '「いいえ/NO」を押した場合は、フォームの回答受付状態/ファイル名/タイトル/回答URLがリストに反映されます。\n' +
        '「キャンセル/CANCEL」を押した場合は、モードが「一時停止」に戻ります。',
        ui.ButtonSet.YES_NO_CANCEL);
      if (confirmation === ui.Button.YES) {
        for (const form of forms) {
          form.setAcceptingResponses(datas[i][0]);
          formFile.setName(datas[i][1]);
          form.setTitle(datas[i][2]);
        }
      } else if (confirmation === ui.Button.NO) {
        flag = true;
      } else {
        editedRange.setValue(e.oldValue);
      }
    }

    if (e.value === 'フォームと同期' || flag) {
      let writeDatas = new Array();

      for (let i = 0; i < forms.length; i++) {
        let contents = new Array();

        contents.push(forms[i].isAcceptingResponses());
        contents.push(formFiles[i].getName());
        contents.push(forms[i].getTitle());
        contents.push(forms[i].getEditUrl());
        const publishedUrl = forms[i].getPublishedUrl();
          contents.push(forms[i].shortenFormUrl(publishedUrl));

        writeDatas.push(contents);
      }

      if (writeDatas[0]) {
        sheet.getRange(6,2,writeDatas.length,writeDatas[0].length).setValues(writeDatas);
      }
      if (e.value === 'フォームと同期') {
        editedRange.setValue(e.oldValue);
      }
    }
  } catch(error) {
    ss.toast(error.message,'エラー');
    console.log('Error: ' + error.message);
  } finally {
    ui.showModalDialog(closeModalDialog, '処理が完了しました');
  }
}

function startOrStopResponse() { //毎分実行
  try {
    const firstRow = 6;
    const rowLength = lastRow - 5;
    if (rowLength < 1 || triggerMode === '一時停止') return;

    const isAcceptingResponses = sheet.getRange(firstRow, 2, rowLength, 1).getValues().flat();
    const startResponseDates = sheet.getRange(firstRow,7,rowLength,1).getValues().flat();
    const stopResponseDates = sheet.getRange(firstRow,8,rowLength,1).getValues().flat();
    const formUrls = sheet.getRange(firstRow,5,rowLength,1).getValues().flat();
    const now = new Date();
    
    for (let i = 0; i < rowLength; i++) {
      let acceptingResponses = null;

      if (startResponseDates[i] && startResponseDates[i] <= now) acceptingResponses = true;
      if (stopResponseDates[i] && stopResponseDates[i] <= now) acceptingResponses = false;

      console.log('acceptingResponses:' + acceptingResponses);

      if (acceptingResponses !== null && acceptingResponses !== isAcceptingResponses) {
        const form = FormApp.openByUrl(formUrls[i]);
        form.setAcceptingResponses(acceptingResponses);
        sheet.getRange(firstRow+i,2).setValue(acceptingResponses);
      }
    }
  } catch(error) {
    console.log('Error: ' + error.message);
  }
}