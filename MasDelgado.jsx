/* =========================================================================

NOMBRE: MasDelgado.jsx
VERSÍON: 1.7

ESCRITOR: Vitaly Batushev
FECHA: 29.09.2012

SOPORTE: http://szam5.com
CORREO: vitaly@batushev.info

============================================================================ */
var strSelectFile = 'Выберите файл для улучшения';
var strQuestionRepair = 'Улучшить активную публикацию?';
var strQuestionOpen = 'Открыть улучшенную публикацию?';
 if (app.documents.length == 0) {
	var mySourceFile = File.openDialog(strSelectFile, '*.INDD');
	if (mySourceFile == null) { exit(); }
} else {
	var rst = confirm(strQuestionRepair);
	if (rst == true) { 
		var mySourceFile = new File(app.activeDocument.fullName);
		app.activeDocument.close(1852776480);
	} else {
		var mySourceFile = File.openDialog(strSelectFile, '*.INDD');
		if (mySourceFile == null) { exit(); }
	}
}
var tmpFolder = Folder(mySourceFile.path);
var idmlFile = new File(tmpFolder + '/temp.idml');
var myName = mySourceFile.fsName;
var bakName = myName.substr(0, myName.length - 5) + '-bak.indd';
var myDoc = app.open(mySourceFile, false);
myDoc.exportFile(ExportFormat.INDESIGN_MARKUP, idmlFile);
myDoc.close(1852776480);
mySourceFile.rename(bakName);
var newDoc = app.open(idmlFile, false);
newDoc.close(2036691744, myName);
idmlFile.remove();
if (confirm(strQuestionOpen)) { app.open(myName, true); }
