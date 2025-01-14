const canDoparsonal = () => {
    
    function getSheet(id){
        // スプレッドシートのID
        const spreadsheetId = id;

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
        return data;

    };

    const mData = getSheet('1mY7UeObWpriZX5DhdFe56p7BcVdTqFzYHD9QGQRDjxs');
    const fData = getSheet('1V5vAiDbUk683CG1ukk1TG8UbngApHdCXHuFqvPoXt5k'); 
    // ひな形のドキュメント (Fileクラスのオブジェクトを取得)
    const templateFile = DriveApp.getFileById('1NSGZf5ddeTJpguTmwPuv2PDC4aQcP3hAMSw_hl6CRNw');

    function writeDocument(){
    // 1年1ホーム1番の中間の点数を取得
    const mScore = mData[0];
    //1年1ホーム1番の期末の点数を、fDataの配列の中からdocumentDataMiddleのgrade,home,numが同じ配列を取得
    const fScore = fData.find(score => 
        score[1] === mScore[1] && 
        score[2] === mScore[2] && 
        score[3] === mScore[3]
    );//データが見つからなかった場合の処理
    // 新しいドキュメントを作成
    const newDocFile = templateFile.makeCopy(`個人票_${mScore[1]}_${mScore[2]}_${mScore[3]}_${mScore[0]}`);
    // ドキュメントの本文を取得
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const body = newDoc.getBody();

    // ドキュメントに書き込む
    const documentData = {
        grade: mScore[1],
        home: mScore[2],
        num: mScore[3],
        name: mScore[0],
        mGreeting: mScore[4],
        mAppearance: mScore[5],
        mClass: mScore[6],
        mListen: mScore[7],
        mBeautification: mScore[8],
        mTime: mScore[9],
        mPhone: mScore[10],
        mReading: mScore[11],
        mAverage: mScore[12],
        fGreeting: fScore[4],
        fAppearance: fScore[5],
        fClass: fScore[6],
        fListen: fScore[7],
        fBeautification: fScore[8],
        fTime: fScore[9],
        fPhone: fScore[10],
        fReading: fScore[11],
        fAverage: fScore[12],
    };
    //ドキュメントに書き込むのに使う配列
    const dDArray = ["grade", "home", "num", "name", "mGreeting", "mAppearance", "mClass", "mListen", "mBeautification", "mTime", "mPhone", "mReading", "mAverage", "fGreeting", "fAppearance", "fClass", "fListen", "fBeautification", "fTime", "fPhone", "fReading", "fAverage"];

    // プレースホルダーを置き換える
    for (let i = 0; i < dDArray.length; i++) {
    body.replaceText(`{{${dDArray[i]}}}`, documentData[dDArray[i]]);
    }
    }

    writeDocument();
    

};