/* =========================================================================

InDesign ExtendScript

NAME: InCountChars
VERSION: 1.2

AUTHOR: Vitaly Batushev
DATE  : 22.12.2013

DESCRIPTION: Статистика о количестве знаков в публикации

E-MAIL: vitaly@batushev.info

============================================================================ */

// Список исключаемых символов по умолчанию.
// Пишется в формате GREP. Можно использовать GREP-символы Adobe InDesign.
// При перечислении символы разделяются символом |
var excludeSymbols = "~-";


with (app) {
var ScriptName = getFileNameOnly(app.activeScript.name) + ' 1.2';
try {		
	if (app.version.substr(0,1) > 3 ) {
		$.locale = "ru";
		//$.localize = true;
		var strMinWords = localize({en: 'Minimum words in Story', ru: 'Минимальное количество слов в статье:'});
		var strViewWithOutSpaces = localize({en: 'Show numbers characters without spaces', ru: 'Показывать количество знаков без пробелов'});
		var strViewWithOutSpecial = localize({en: 'Show numbers characters without special symbols', ru: 'Показывать количество знаков без указанных символов:'});
		var strAttention = localize({en: 'Attention!', ru: 'Внимание!'});
		var strNoOpenDocument = localize({en: 'No open documents!', ru: 'Нет открытых документов!'});
		var strReport = localize({en: 'Report', ru: 'Статистический отчет'});
		var strPublication = localize({en: 'Publication', ru: 'Название публикации'});
		var strCharsWithSpaces = localize({en: 'characters with spaces', ru: 'символов, включая пробелы'});
		var strCharsWithoutSpaces = localize({en: 'characters without spaces', ru: 'символов, не включая пробелы'});
		var strCharsWithoutSymbols = localize({en: 'characters exclude', ru: 'символов, не включая'});
		var strAuthorLists = localize({en: 'Authors lists', ru: 'Авторских листов'});
		var strLessOneList = localize({en: 'less than one authors list', ru: 'менее одного авторского листа'});
		var strSaveReport = localize({en: 'Save report in file?', ru: 'Сохранить отчет в файл?'});
		var strTextFiles = localize({en: 'Text Files', ru: 'Текстовые файлы'});
		var strSelectResultFile = localize({en: 'Save Result File', ru: 'Сохранить результат'});
		var strEndWork = localize({en: 'Report created!', ru: 'Создание отчета завершено!'});
		var strNoSaveFile = localize({en: 'File is not saved', ru: 'Файл не сохранен!'});
		var strFileSaveAs = localize({en: 'File save as', ru: 'Файл сохранен как'});
		var strCreatedWith = localize({en: 'Created with script', ru: 'Создано с помощью скрипта'});
		var strCredits = localize({en: 'by Vitaly Batushev (https://github.com/vbatushev/)', ru: 'Виталий Батушев (https://github.com/vbatushev/)'});
		var strBars = localize({en: '===============================', ru: '==============================='});
	} else {
		var strMinWords = 'Minimum words in Story';
		var strViewWithOutSpaces = 'Show numbers characters without spaces';
		var strViewWithOutSpecial = 'Show numbers characters without special symbols';
		var strAttention = 'Attention!';
		var strNoOpenDocument = 'No open documents!';
		var strReport = 'Report';
		var strPublication = 'Publication';
		var strCharsWithSpaces = 'characters with spaces';
		var strCharsWithoutSpaces = 'characters without spaces';
		var strCharsWithoutSymbols = 'characters exclude';
		var strAuthorLists = 'Authors lists';
		var strLessOneList = 'less than one authors list';
		var strSaveReport = 'Save report in file?';
		var strTextFiles = 'Text Files';
		var strSelectResultFile = 'Save Report File';
		var strEndWork = 'Report created!';
		var strNoSaveFile = 'File is not saved';
		var strFileSaveAs = 'File save as';
		var strCreatedWith = 'Created with script';
		var strCredits = 'by Vitaly Batushev (http://indesing.rudtp.ru)';
		var strBars = '===============================';
	}
} catch (e) {
	var strMinWords = 'Minimum words in Story';
	var strViewWithOutSpaces = 'Show numbers characters without spaces';
	var strViewWithOutSpecial = 'Show numbers characters without special symbols';
	var strCharsWithoutSymbols = 'characters exclude';
	var strAttention = 'Attention!';
	var strNoOpenDocument = 'No open documents!';
	var strReport = 'Report';
	var strPublication = 'Publication';
	var strCharsWithSpaces = 'characters with spaces';
	var strCharsWithoutSpaces = 'characters without spaces';
	var strAuthorLists = 'Authors lists';
	var strLessOneList = 'less than one authors list';
	var strSaveReport = 'Save report in file?';
	var strTextFiles = 'Text Files';
	var strSelectResultFile = 'Save Report File';
	var strEndWork = 'Report created!';
	var strNoSaveFile = 'File is not saved';
	var strFileSaveAs = 'File save as';
	var strCreatedWith = 'Created with script';
	var strCredits = 'by Vitaly Batushev (http://szam5.com)';
	var strBars = '===============================';
}
	if (documents.length < 1) { alert(strNoOpenDocument, strAttention); exit(); }
	var myDoc = activeDocument;	
	var myDlg = dialogs.add({name:ScriptName});
	with(myDlg) {
		with(dialogColumns.add()) {
			with (dialogRows.add()) {
				staticTexts.add({staticLabel:strMinWords});
				var minWords = integerEditboxes.add({editValue:0, minWidth:25});
			}
			with (dialogRows.add()) {
				staticTexts.add({staticLabel:strViewWithOutSpecial});
			}
			with (dialogRows.add()) {
				var outSmbl = textEditboxes.add({editContents:excludeSymbols, minWidth:340});
			}
			var chkWithOutSpace = checkboxControls.add({staticLabel:strViewWithOutSpaces, checkedState:false})
		}
	}
	result = myDlg.show();
	if (result == false) { myDlg.destroy(); exit();	}
	var fullCount = 0;
	var spaceCount = 0;
	var outSymbols = 0;
	for (i = 0; i < myDoc.stories.length; i++) {
		var myStory = myDoc.stories[i];
		if (myStory.words.length > minWords.editValue) {
			fullCount = fullCount + myStory.characters.length;
			if (chkWithOutSpace.checkedState == true ){
				var  re = /\s/g;
				try {
					var tmpSpace = myStory.contents.match(re);
				} catch (e) {}
				if (tmpSpace != null) { spaceCount = spaceCount + tmpSpace.length; }
			}
			if (outSmbl.editContents != "") {
				findGrepPreferences = null;
				findGrepPreferences = null;
				findGrepPreferences.findWhat = outSmbl.editContents;
				var findSymbols = myStory.findGrep();
				outSymbols += findSymbols.length;				
				changeGrepPreferences = null;
			}
			if (myStory.tables.length > 0) {
				for (ti = 0; ti < myStory.tables.length; ti++) {
					var myTable = myStory.tables[ti];
					for (j = 0; j < myTable.cells.length; j++) {
						fullCount = fullCount + myTable.cells[j].characters.length;
						if (chkWithOutSpace.checkedState == true ){ 
							try {
								var tmpSpace = myTable.cells[j].contents.match(re);
							} catch (e) {}
							if (tmpSpace != null) { spaceCount = spaceCount + tmpSpace.length; }
						}
					}
				}
			}
		}
	}
	var aList = Math.round(fullCount/40000);
	if (aList < 1) { aList = strLessOneList };
	var txtReport = strBars + '\n' + strReport + '\n' + strBars + '\n\n' + 
	strPublication + ': ' + activeDocument.name + '\n\n' +
	fullCount + ' ' + strCharsWithSpaces + '.' + '\n';
	if (chkWithOutSpace.checkedState == true ){
		txtReport += (fullCount - spaceCount) + ' ' + strCharsWithoutSpaces + '.' + '\n\n';
	}
	if (outSymbols > 0) {
		txtReport += (fullCount - outSymbols) + ' ' + strCharsWithoutSymbols + " " + outSmbl.editContents + '.' + '\n\n';
	}
	txtReport = txtReport + strAuthorLists + ': ' + aList + '.' + '\n\n\n\n' + strBars + '\n' + 
	strCreatedWith + ' ' + ScriptName + '\n' + strCredits;
	alert (txtReport, ScriptName);

	var myResultFile = new File (activeDocument.filePath + "/Report.txt");
	myResultFile.encoding = 'UTF16';	
	myResultFile.open("w");		
	myResultFile.write(txtReport);
	myResultFile.close();	
	myResultFile.execute();
	
}

function getFileNameOnly (myFileName) {
	var myString = "";
	var myResult = myFileName.lastIndexOf(".");
	if (myResult == -1) { myString = myFileName; }
	else { myString = myFileName.substr(0, myResult); }
	return myString;
}
