const canDoParsonal = () => {
    
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
    //ドキュメントの保存先を指定
    // 保存先のフォルダIDを取得
    const folderId = "1wpBjCqE76hCihZp7xj2iosoY_DaX5lEg"; 
    const folder = DriveApp.getFolderById(folderId);

    //得点の推移を計算する関数
    function calculateTransition(mScore, fScore) {
        if (fScore === "未提出" || mScore === "未提出") {
            return "---";
        } else {
            if (mScore > fScore) {
                return "-" + (mScore - fScore);
            } else if (mScore < fScore) {
                return "+" + (fScore - mScore);
            } else {
                return "±0";
            }
        }
    }

    // ドキュメントに書き込む関数
    function writing(mScore, fScore) {
        // 新しいドキュメントを作成
        const newDocFile = templateFile.makeCopy(`個人票_${mScore[1]}_${mScore[2]}_${mScore[3]}_${mScore[0]}`);
        newDocFile.moveTo(folder);

        // ドキュメントの本文を取得
        const newDoc = DocumentApp.openById(newDocFile.getId());
        const body = newDoc.getBody();
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
            tGreeting: calculateTransition(mScore[4], fScore[4]),
            tAppearance: calculateTransition(mScore[5], fScore[5]),
            tClass: calculateTransition(mScore[6], fScore[6]),
            tListen: calculateTransition(mScore[7], fScore[7]),
            tBeautification: calculateTransition(mScore[8], fScore[8]),
            tTime: calculateTransition(mScore[9], fScore[9]),
            tPhone: calculateTransition(mScore[10], fScore[10]),
            tReading: calculateTransition(mScore[11], fScore[11]),
            tAverage: calculateTransition(mScore[12], fScore[12])
        };
        //ドキュメントに書き込むのに使う配列
        const dDArray = ["grade", "home", "num", "name", "mGreeting", "mAppearance", "mClass", "mListen", "mBeautification", "mTime", "mPhone", "mReading", "mAverage", "fGreeting", "fAppearance", "fClass", "fListen", "fBeautification", "fTime", "fPhone", "fReading", "fAverage", "tGreeting", "tAppearance", "tClass", "tListen", "tBeautification", "tTime", "tPhone", "tReading", "tAverage"];

        // プレースホルダーを置き換える
        for (let i = 0; i < dDArray.length; i++) {
        body.replaceText(`{{${dDArray[i]}}}`, documentData[dDArray[i]]);
        }
    }

//mScoreが存在する生徒の数だけドキュメントを作成
for (let i = 0; i < 3; i++) {
    // 1年1ホームi番の中間の点数を取得
    const mScore = mData[i];
    //1年1ホーム1番の期末の点数を、fDataの配列の中からdocumentDataMiddleのgrade,home,numが同じ配列を取得
    let fScore = fData.find(score => 
        score[1] === mScore[1] && 
        score[2] === mScore[2] && 
        score[3] === mScore[3]
    );
    //データが見つからなかった場合の処理   
    if (!fScore) {
        fScore = [mScore[0], mScore[1], mScore[2], mScore[3], "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出"];
    }
    //fDataからmScoreと同じ配列を削除
    const index = fData.indexOf(fScore);
    if (index > -1) {
    fData.splice(index, 1);
    }
    writing(mScore, fScore);
}
//fDataのみ存在する生徒の数だけドキュメントを作成 error
if (fData.length > 0) {
    for (let i = 0; i < fData.length; i++) {
        const fScore = fData[i];
        const mScore = [fScore[0],fScore[1], fScore[2], fScore[3], "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出"];
        writing(mScore, fScore);
    }
}else{
    console.log("fData is empty");
}
};