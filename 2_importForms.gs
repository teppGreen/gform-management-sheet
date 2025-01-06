function listForms() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = ui.prompt('フォームの編集URL または フォルダの共有URL を入力してください', 
      '● フォルダの場合は、その階層にある全てのフォーム（ショートカットを含む）が追加されます（下層フォルダは含まれません）\n' + 
      '● 一度に入力できるURLは1つまでです\n' + 
      '● 入力するフォルダ・フォームの編集権限があなたにない場合はエラーになります', 
    ui.ButtonSet.OK_CANCEL);
      if (url.getSelectedButton() === ui.Button.CANCEL) {
        throw new Error('キャンセルボタンが押されたため、処理を中断しました。')
      }
    const id = extractFileId(url.getResponseText());

    if (!id) {
      throw new TypeError('ファイル/フォルダが見つかりませんでした。正しいURLを入力してください。');
    }

    const mimeType = DriveApp.getFileById(id).getMimeType();
    let formIds = new Array();

    if (mimeType === MimeType.FOLDER) {
      const folder = DriveApp.getFolderById(id);
      const folderName = folder.getName();

      loadingAnimation.message = 'フォルダ「' + folderName + '」の中を検索しています。';
      ui.showModalDialog(loadingAnimation.evaluate().setWidth(400).setHeight(320), "処理中");

      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() === MimeType.GOOGLE_FORMS) {
          formIds.push(file.getId())
        } else if (file.getTargetMimeType() === MimeType.GOOGLE_FORMS) {
          formIds.push(file.getTargetId());
        }
      }

      if (formIds.length === 0) {
        throw new TypeError('フォルダの中にフォームがありませんでした。');
      }
    } else if(mimeType === MimeType.GOOGLE_FORMS){
      formIds.push(id);
    } else {
      throw new TypeError('入力されたURLはフォームの編集URLではありません。');
    }

    let writeDatas = new Array();
    const existForms = sheet.getRange(1,5,lastRow,1).getValues().flat();

    for (const formId of formIds) {
      const form = FormApp.openById(formId);
      
      let contents = new Array();
      contents.push(form.isAcceptingResponses()); //contents[0]
      contents.push(DriveApp.getFileById(formId).getName()); //contents[1]
      contents.push(form.getTitle()); //contents[2]
      const editUrl = form.getEditUrl();
        contents.push(editUrl); //contents[3]
      const publishedUrl = form.getPublishedUrl();
        contents.push(form.shortenFormUrl(publishedUrl)); //contents[4]

      if (existForms.includes(contents[3])) {
        const isOverLength = contents[1].length > 24 ? '…' : '';
        throw new Error('「' + contents[1].substr(1,24) + isOverLength + '」は既に追加されています。');
      } else {
        writeDatas.push(contents);
      }
    }
    
    console.log(writeDatas);
    if (writeDatas[0]) {
      sheet.getRange(lastRow + 1, 2, writeDatas.length, writeDatas[0].length).setValues(writeDatas);
    }
  } catch(error) {
    ss.toast(error.message, 'エラー');
    console.log('Error: ' + error.message);
  } finally {
    //ローディングアニメーションを閉じる
    ui.showModalDialog(closeModalDialog, '処理が終了しました');
  }
}

function clearList() {
  const rowLength = lastRow-5;
  if (rowLength > 0) {
    sheet.getRange(6,2,rowLength,8).clearContent();
  }
}