/*
* @Author: Vitaly Batushev
* @E-Mail: vitaly@batushev.info
* @Date:   2015-12-07 02:43:24
* @Last Modified by:   Vitaly Batushev
* @Last Modified time: 2016-03-19 22:08:54
* @ScriptName  RecoveryINDD
* @Version     1.8
* @Description Скрипт для "лечения" или уменьшения размера публикации Adobe InDesign.
*/

var RecoveryINDD = (function(){
    var strings = {
        select: "Выберите файл для исправлени",
        questions: {
            repair: "Исправить активную публикацию?",
            open: "Открыть улучшенную публикацию?",
        }
    }

    var main = function() {
        var mySourceFile;
         if (app.documents.length == 0) {
            mySourceFile = File.openDialog(strings.select, "*.INDD");
            if (mySourceFile === null) {
                exit();
            }
        } else {
            var rst = confirm(strings.questions.repair, false, "Recovery INDD");
            if (rst == true) {
                mySourceFile = new File(app.activeDocument.fullName);
                app.activeDocument.close(SaveOptions.NO);
            } else {
                var mySourceFile = File.openDialog(strings.select, "*.INDD");
                if (mySourceFile === null) {
                    exit();
                }
            }
        }

        var idmlFile = new File(mySourceFile.path + "/temp.idml");
        var backName = File(mySourceFile.absoluteURI + ".bak");
        var myDoc = app.open(mySourceFile, false);

        myDoc.exportFile(ExportFormat.INDESIGN_MARKUP, idmlFile);
        myDoc.close(SaveOptions.NO);

        mySourceFile.copy(backName);

        var newDoc = app.open(idmlFile, false);
        newDoc.close(SaveOptions.YES, mySourceFile);
        idmlFile.remove();

        if (confirm(strings.questions.open, false, "Recovery INDD")) {
            app.open(mySourceFile, true);
        }
    }

    return {
        run: main
    }
})();
RecoveryINDD.run();