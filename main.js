const canDoParsonal = () => {
    
    function getSheetData(spreadsheetIds, sheetName) {
        const data = {};
        spreadsheetIds.forEach((spreadsheetId, index) => {
            const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
            const sheet = spreadsheet.getSheetByName(sheetName);
            const lastRow = sheet.getLastRow();
            const lastColumn = sheet.getLastColumn();
            data[index] = sheet.getRange(2, 2, lastRow - 1, lastColumn).getValues();
        });
        return data;
    }
    
    const spreadsheetIds = ['1mY7UeObWpriZX5DhdFe56p7BcVdTqFzYHD9QGQRDjxs', '1V5vAiDbUk683CG1ukk1TG8UbngApHdCXHuFqvPoXt5k'];
    const sheetData = getSheetData(spreadsheetIds, '各学年の点数'); // シート名は同じと仮定
    const mData = sheetData[0];
    const fData = sheetData[1];
    
    
    // ひな形のドキュメント
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

        // ドキュメントの本文を取得
        // ドキュメントのコピーを作成し、Document オブジェクトを取得
        const newDoc = templateFile.makeCopy(`個人票_${mScore[1]}_${mScore[2]}_${mScore[3]}_${mScore[0]}`, folder).getAs('application/vnd.google-apps.document'); 
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
        // 一度に置き換え
        body.replaceText("{{grade}}", documentData.grade)
            .replaceText("{{home}}", documentData.home)
            .replaceText("{{num}}", documentData.num)
            .replaceText("{{name}}", documentData.name)
            .replaceText("{{mGreeting}}", documentData.mGreeting)
            .replaceText("{{mAppearance}}", documentData.mAppearance)
            .replaceText("{{mClass}}", documentData.mClass)
            .replaceText("{{mListen}}", documentData.mListen)
            .replaceText("{{mBeautification}}", documentData.mBeautification)
            .replaceText("{{mTime}}", documentData.mTime)
            .replaceText("{{mPhone}}", documentData.mPhone)
            .replaceText("{{mReading}}", documentData.mReading)
            .replaceText("{{mAverage}}", documentData.mAverage)
            .replaceText("{{fGreeting}}", documentData.fGreeting)
            .replaceText("{{fAppearance}}", documentData.fAppearance)
            .replaceText("{{fClass}}", documentData.fClass)
            .replaceText("{{fListen}}", documentData.fListen)
            .replaceText("{{fBeautification}}", documentData.fBeautification)
            .replaceText("{{fTime}}", documentData.fTime)
            .replaceText("{{fPhone}}", documentData.fPhone)
            .replaceText("{{fReading}}", documentData.fReading)
            .replaceText("{{fAverage}}", documentData.fAverage)
            .replaceText("{{tGreeting}}", documentData.tGreeting)
            .replaceText("{{tAppearance}}", documentData.tAppearance)
            .replaceText("{{tClass}}", documentData.tClass)
            .replaceText("{{tListen}}", documentData.tListen)
            .replaceText("{{tBeautification}}", documentData.tBeautification)
            .replaceText("{{tTime}}", documentData.tTime)
            .replaceText("{{tPhone}}", documentData.tPhone)
            .replaceText("{{tReading}}", documentData.tReading)
            .replaceText("{{tAverage}}", documentData.tAverage);
    }

    const fDataObject = {};
    fData.forEach(score => {
        const key = `${score[1]}_${score[2]}_${score[3]}`;
        fDataObject[key] = score;
    });

//mScoreが存在する生徒の数だけドキュメントを作成
for (let i = 0; i < mData.length; i++) {
    const mScore = mData[i];
    const key = `${mScore[1]}_${mScore[2]}_${mScore[3]}`;
    let fScore = fDataObject[key] || [mScore[0], mScore[1], mScore[2], mScore[3], "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出", "未提出"];

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