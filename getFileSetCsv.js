//特定フォルダ取得
function getFolder() {
    //フォルダID定義※フォルダのIDを入れる
    var folder_id = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    //特定フォルダ取得
    var my_folder = DriveApp.getFolderById(folder_id);

    return my_folder;
}

//フォルダ内削除
function deleteFile() {
    //フォルダ取得
    var my_folder = getFolder();
    var files = my_folder.getFiles();
    //ファイルオブジェクトを一つずつ展開
    for (var i = 0; files.hasNext(); i++) {
        //ファイルの取得
        var file = files.next();
        //ファイル名とファイルIDの削除
        my_folder.removeFile(file)
    }
}



//添付メール取得➡特定フォルダ取得➡フォルダ内削除➡フォルダに保存
function getMailAttachment() {
    //メール検索ロジックを入れる
    var serch_subject = 'XXXXXXXXXXXXXXXXXXXXXXXX'
    var date = 'after:' + Moment.moment().format('YYYY/MM/DD').toString()
    //上記条件でメールを取得
    var query = serch_subject + date
    var my_threads = GmailApp.search(query);
    //メール自体をスレッドとして二元配列で格納
    var my_messages = GmailApp.getMessagesForThreads(my_threads);

    //フォルダ取得
    var my_folder = getFolder();
    //フォルダ内削除
    deleteFile()
    //二元配列を展開してメール内容から添付ファイルを取得
    for (var i in my_messages) {
        var attachments = my_messages[i].slice(-1)[0].getAttachments(); //添付ファイルを取得
        for (var k in attachments) {
            if (i < 2) {
                //ファイル名でチェックする
                if (attachments[k].getName().indexOf('XXXXXXXXXXXXXXXXXXX') != -1) {
                    attachments[k].setName('XXXXXXXXXXXXXXXXXXXXX');
                    //ドライブに添付ファイルを保存
                    my_folder.createFile(attachments[k]);
                    //ファイル名でチェックする
                } else if (attachments[k].getName().indexOf('XXXXXXXXXXXXXXXXXXXXX') != -1) {
                    attachments[k].setName('XXXXXXXXXXXXXXXXXXXXX');
                    //ドライブに添付ファイルを保存
                    my_folder.createFile(attachments[k]);
                } else {
                    console.log('マッチしませんでした。');
                    break;
                }
            }
        }
    }
}

function setCsvToSpredSheet() {
    //書き込む対象のSpread Sheetを定義
    var spreadsheet = SpreadsheetApp.openById('XXXXXXXXXXXXXXXXXXX');

    //フォルダ取得
    var my_folder = getFolder()
    var files = my_folder.getFiles();
    for (var i = 0; files.hasNext(); i++) {
        var file = files.next();
        //どのフォルダを取得するか決める
        if (file.getName() == 'XXXXXXXXXXXXXXXXXXXX') {
            //書き込むスプレッドシートのシート名を入力
            var sh = spreadsheet.getSheetByName('XXXXXXXXXXXXXXXXXXXX');
            sh.clear();
            //設定しないでShift_JISのままだと文字化けする可能性あり
            var data = file.getBlob().getDataAsString("Shift_JIS");

            var csv = Utilities.parseCsv(data);
            //セルA1からCSVの内容を書き込んでいく
            sh.getRange(1, 1, csv.length, csv[0].length).setValues(csv);
        }
    }
}
