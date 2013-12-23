/* =========================================================================

NAME: InBackup
VERSION: 1.02

AUTHOR: Vitaly Batushev
DATE  : 21.12.2013

DESCRIPTION: Создание резервных копий документа в папке указываемой пользователем

============================================================================ */

/* ====================== Н А С Т Р О Й К И ================================ */
var fldBackup = '~/Documents/Backup';	// Папка для сохранения резервной копии
var manyBackups = true;			
var numBackups = 5;  			// null = бесконечное количество, число – количество резервных копий
var fmtBackups = 'n'; 			// d - дата/время, n - номер
var usePrefix = false; 			// если значение true, используется префикс, например, 00001-myPub.indd
								// если значение false, используется суффикс, например, myPub-00001.indd
var sepPrefix = '-';			// Символ, отделяющий префикс от имени файла
var sepSuffix = '-';			// Символ, отделяющий суффикс от имени файла
var fmtDate = 'D.M.Y-h:m-s';		// D - день, M - месяц, Y - год (4-значное значение),
					// h - часы, m - минуты, s - секунды
					// НЕ ИСПОЛЬЗУЙТЕ СИМВОЛЫ ':', '/', '\'
/* ================= К О Н Е Ц    Н А С Т Р О Е К =========================== */

var bkFolder = new Folder(fldBackup);
if (!bkFolder.exists) { bkFolder.create(); }
if (app.documents.length == 0) { exit(); }
var myDoc = app.activeDocument;
myDoc.save();
var myFile = new File(myDoc.fullName);
var myDocName = getFileNameOnly(myDoc.name);
myFile.copy(generateName(myDocName));
function getFileNameOnly (myFileName) {
	if (myFileName.lastIndexOf(".") == -1) { return myFileName; } else { return myFileName.substr(0, myFileName.lastIndexOf(".")); }
}
function generateName(fileName) {
if (manyBackups) {
	switch (fmtBackups) {
		case 'd':
			var curDate = new Date();
			fmtDate = fmtDate.replace(/D/g, curDate.getDate()).replace(/M/g, curDate.getMonth()).replace(/Y/g, curDate.getFullYear()).replace(/h/g, curDate.getHours()).replace(/m/g, curDate.getMinutes()).replace(/s/g, curDate.getSeconds()).replace(/[\/\?\:,\\\|]/g, '-');
			var mySuffix = sepSuffix + fmtDate;
			if (usePrefix) { myPrefix = sepPrefix + mySuffix; mySuffix = ''; } else { myPrefix = ''; }
			mySuffix += '.indd';
			var myName = fldBackup + '/' + myPrefix + fileName + mySuffix;
			break;
		case 'n':
			if (numBackups == null) {
				var num = 0;
				do {
					num++;
					switch (num.toString().length) {
						case 1: num = '0000' + num; break;
						case 2: num = '000' + num; break;
						case 3:	num = '00' + num; break;
						case 4:	num = '0' + num; break;
						default: break;
					}					
					if (usePrefix) {
						var qFile = new File(fldBackup + '/' + num + sepPrefix + fileName + '.indd');
					} else {
						var qFile = new File(fldBackup + '/' + fileName + sepSuffix + num  + '.indd');
					}
				} while (qFile.exists);
				var myName = qFile.absoluteURI;
			} else {
				var num = app.activeDocument.extractLabel('cntBackup');
				if (num == '') { num = 1; }
				if (num == numBackups) {
					app.activeDocument.insertLabel('cntBackup', '1');
				} else {
					app.activeDocument.insertLabel('cntBackup', (Math.round(num) + 1).toString());				
				}
				if (usePrefix) {
					var myName = new File(fldBackup + '/' + num + sepPrefix + fileName + '.indd');
				} else {
					var myName = new File(fldBackup + '/' + fileName + sepSuffix + num  + '.indd');
				}
			}
			break;
		default:
			var myName = fileName + '.indd';
		}						
} else {
	var myName = fileName + '.indd';
}
	return myName;
}
