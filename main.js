const canDoparsonal = () => {
    // スプレッドシートのID
    const spreadsheetId = '1mY7UeObWpriZX5DhdFe56p7BcVdTqFzYHD9QGQRDjxs';

    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    // シートを取得
    const sheet = spreadsheet.getSheetByName('各学年の点数');

    // 最終行を取得
    const lastRow = sheet.getLastRow();

    // 最終列を取得
    const lastColumn = sheet.getLastColumn();

    // データを取得
    const range = sheet.getRange(2, 2, lastRow, lastColumn);
    const data = range.getValues();

    // ひな形のドキュメント (Fileクラスのオブジェクトを取得)
    const templateFile = DriveApp.getFileById('1NSGZf5ddeTJpguTmwPuv2PDC4aQcP3hAMSw_hl6CRNw');

    // 新しいドキュメントを作成
    const newDocFile = templateFile.makeCopy('新しいドキュメント');

    // ドキュメントの本文を取得
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const body = newDoc.getBody();

    // 1年1ホーム1番の点数を取得
    const score = data[0];

    // ドキュメントに書き込む
    const documentData = {
        grade: score[1],
        home: score[2],
        num: score[3],
        name: score[0],
    };

    // プレースホルダーを置き換える
    body.replaceText('{{grade}}', documentData.grade);
    body.replaceText('{{home}}', documentData.home);
    body.replaceText('{{num}}', documentData.num);
    body.replaceText('{{name}}', documentData.name);
};