/*
* @Author: Vitaly Batushev
* @E-Mail: vitaly@batushev.info
* @Date:   2015-12-07 02:43:24
* @Last Modified by:   Vitaly Batushev
* @Last Modified time: 2016-03-19 22:08:54
* @ScriptName  InBackup
* @Version     1.1
* @Description Создание резервных копий документа в папке указываемой пользователем
*/

/* ====================== Н А С Т Р О Й К И ================================ */

var InBackupConfig = {
    folder: "~/Documents/Backup",   // Папка для сохранения резервной копии
    many: true,
    num: 5,                         // null = бесконечное количество, число – количество резервных копий
    format: "d",                   // d - дата/время, n - номер
    usePrefix: false,               // если значение true, используется префикс, например, 00001-myPub.indd
                                    // если значение false, используется суффикс, например, myPub-00001.indd
    prefix: "-",                    // Символ, отделяющий префикс от имени файла
    suffix: "-",                    // Символ, отделяющий суффикс от имени файла
    dateformat: "D.M.Y-h:m-s"      // D - день, M - месяц, Y - год (4-значное значение),
                                    // h - часы, m - минуты, s - секунды
                                    // НЕ ИСПОЛЬЗУЙТЕ СИМВОЛЫ ':', '/', '\'
}
/* ================= К О Н Е Ц    Н А С Т Р О Е К =========================== */

var InBackup = (function(){
    var main = function() {
        if (app.documents.length == 0) { exit(); }
        var bkFolder = new Folder(InBackupConfig.folder);
        if (!bkFolder.exists) {
            reverseFolderCreate(bkFolder);
        }
        app.activeDocument.save();
        var activeFile = new File(app.activeDocument.fullName);
        var newFile = generateName(getFileNameOnly(app.activeDocument.name));
        activeFile.copy(newFile);
    }

    var getFileNameOnly = function(str) {
        var arr = str.split(".");
        if (arr.length > 1) {
            arr.pop();
        }
        return arr.join(".");
    }

    var reverseFolderCreate = function(fld) {
        var fldpath = fld.absoluteURI;
            paths = fldpath.split("/"),
            rel = paths[0] + "/" + paths[1];
        for (var a = 2, l = paths.length; a < l; a++) {
            rel += "/" + paths[a];
            if (!Folder(rel).exists) {
                Folder(rel).create();
            }
        }
    }

    var generateName = function(fileName) {
        var name = "", num = 0;
        if (InBackupConfig.many) {
            switch (InBackupConfig.format) {
                case "d":
                    var curDate = new Date();
                    fmtDate = InBackupConfig.dateformat
                            .replace(/D/g, curDate.getDate())
                            .replace(/M/g, curDate.getMonth())
                            .replace(/Y/g, curDate.getFullYear())
                            .replace(/h/g, curDate.getHours())
                            .replace(/m/g, curDate.getMinutes())
                            .replace(/s/g, curDate.getSeconds())
                            .replace(/[\/\?\:,\\\|]/g, "-");
                    var fileSuffix = InBackupConfig.suffix + fmtDate;
                    if (InBackupConfig.usePrefix) {
                        myPrefix = InBackupConfig.prefix + fileSuffix;
                        fileSuffix = "";
                    } else {
                        myPrefix = "";
                    }
                    fileSuffix += ".indd";
                    name = InBackupConfig.folder + "/" + myPrefix + fileName + fileSuffix;
                    break;
                case "n":
                    if (InBackupConfig.num == null) {
                        do {
                            num++;
                            num = Array(Math.abs(5 - num.toString().length) + 1).join("0") + "" + num;
                            var qFile = new File(InBackupConfig.folder + "/" + fileName + InBackupConfig.suffix + num  + ".indd");
                            if (usePrefix) {
                                qFile = new File(InBackupConfig.folder + "/" + num + InBackupConfig.prefix + fileName + ".indd");
                            }
                        } while (qFile.exists);
                        name = qFile.absoluteURI;
                    } else {
                        num = app.activeDocument.extractLabel("countBackup");
                        if (num == "") {
                            num = 1;
                        }
                        if (num == InBackupConfig.num) {
                            app.activeDocument.insertLabel("countBackup", "1");
                        } else {
                            app.activeDocument.insertLabel("countBackup", (Math.round(num) + 1).toString());
                        }
                        name = new File(InBackupConfig.folder + "/" + fileName + InBackupConfig.suffix + num  + ".indd");
                        if (InBackupConfig.usePrefix) {
                            name = new File(InBackupConfig.folder + "/" + num + InBackupConfig.prefix + fileName + ".indd");
                        }
                    }
                    break;
                default:
                    name = fileName + ".indd";
                }
        } else {
            name = fileName + ".indd";
        }
            return name;
    }

    return {
        run: main
    }
})();
InBackup.run();