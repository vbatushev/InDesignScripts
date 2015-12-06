/*
* @Author: vbatushev
* @Date:   2015-09-01 09:25:06
* @version:    2.0
* @Last Modified by:   vbatushev
* @Last Modified time: 2015-12-07 02:40:49
*/

// Список исключаемых символов по умолчанию.
// Пишется в формате GREP. Можно использовать GREP-символы Adobe InDesign.
// При перечислении символы разделяются символом |
var excludeSymbols = "~-";

app.scriptPreferences.enableRedraw = true;
var strings = {
    name: 'inCountChars 2.0',
    minWords: 'Minimum words in Story',
    viewWithOutSpaces: 'Show numbers characters without spaces',
    viewWithOutSpecial: 'Show numbers characters without special symbols',
    attention: 'Attention!',
    noOpenDocument: 'No open documents!',
    report: 'Report',
    publication: 'Publication',
    charsWithSpaces: 'characters with spaces',
    charsWithoutSpaces: 'characters without spaces',
    charsWithoutSymbols: 'characters exclude',
    authorLists: 'Authors lists',
    lessOneList: 'less than one authors list',
    saveReport: 'Save report in file?',
    textFiles: 'Text Files',
    selectResultFile: 'Save Report File',
    endWork: 'Report created!',
    noSaveFile: 'File is not saved',
    fileSaveAs: 'File save as',
    createdWith: 'Created with script',
    credits: 'by Vitaly Batushev (https://github.com/vbatushev/InDesignScripts)',
    bars: '===============================',
}

if (app.documents.length < 1) { alert(strings.noOpenDocument, strings.attention); exit(); }

var dlg = app.dialogs.add({name:strings.name});
with(dlg) {
    with(dialogColumns.add()) {
        with (dialogRows.add()) {
            staticTexts.add({staticLabel:strings.minWords});
            var minWords = integerEditboxes.add({editValue:0, minWidth:25});
        }
        with (dialogRows.add()) {
            staticTexts.add({staticLabel:strings.viewWithOutSpecial});
        }
        with (dialogRows.add()) {
            var outSmbl = textEditboxes.add({editContents:excludeSymbols, minWidth:340});
        }
        var chkWithOutSpace = checkboxControls.add({staticLabel:strings.viewWithOutSpaces, checkedState:false})
    }
}
result = dlg.show();
if (result == false) { dlg.destroy(); exit(); }
var counter = new getCountChars();
counter.exclude = outSmbl.editContents;
counter.spaces = chkWithOutSpace.checkedState;
counter.min = minWords.editValue;
counter.run();

function getCountChars() {
    this.exclude = excludeSymbols;
    this.min = 0;
    this.spaces = false;
    this.progress;

    this.run = function() {
        this.progress = new ProgressbarClass(app.activeDocument.stories.length, strings.name);
        var chars = {spaces: 0, symbols: 0, total: 0 }
        for (i = 0; i < app.activeDocument.stories.length; i++) {
            var myStory = app.activeDocument.stories[i];
            if (myStory.words.length > this.min) {
                var symbols = this.getChars(myStory, false);
                chars.total += symbols.total;
                chars.spaces += symbols.spaces;
                chars.symbols += symbols.symbols;

                for (ti = 0; ti < myStory.tables.length; ti++) {
                    var myTable = myStory.tables[ti];
                    for (j = 0; j < myTable.cells.length; j++) {
                        var tblSymbols = this.getChars(myTable.cells[j],false);
                        chars.total += tblSymbols.total;
                        chars.spaces += tblSymbols.spaces;
                        chars.symbols += tblSymbols.symbols;
                    }
                }

                for (f = 0; f < myStory.footnotes.length; f++) {
                    var fSymbols = this.getChars(myStory.footnotes[f],true);
                    chars.total += fSymbols.total;
                    chars.spaces += fSymbols.spaces;
                    chars.symbols += fSymbols.symbols;
                }
            }
            this.progress.increase();
        }
        this.save(chars);
        this.progress.close();
    }

    this.save = function (result) {
        var aList = Math.round(result.total/40000);
        if (aList < 1) { aList = strings.lessOneList };
        var txtReport = strings.bars + '\n' + strings.report + '\n' + strings.bars + '\n\n' +
        strings.publication + ': ' + app.activeDocument.name + '\n\n' +
        result.total + ' ' + strings.charsWithSpaces + '.' + '\n';
        if (this.spaces){
            txtReport += (result.total - result.spaces) + ' ' + strings.charsWithoutSpaces + '.' + '\n\n';
        }
        if (result.symbols > 0) {
            txtReport += (result.total - result.symbols) + ' ' + strings.charsWithoutSymbols + " " + this.exclude + '.' + '\n\n';
        }

        txtReport = txtReport + strings.authorLists + ': ' + aList + '.' + '\n\n\n\n' + strings.bars + '\n' +
        strings.createdWith + ' ' + strings.name + '\n' + strings.credits;
        alert (txtReport, strings.name);

        var myResultFile = new File (app.activeDocument.filePath + "/Otchet.txt");
        myResultFile.encoding = 'UTF8';
        myResultFile.open("w");
        myResultFile.write(txtReport);
        myResultFile.close();
        myResultFile.execute();
    }

    this.getChars = function(obj, footnote) {
        var out = new Object();
        var content = obj.contents;
        out.total = obj.characters.length;
        var  re = /\s/g;
        var tmpSpace = content.match(re);
        if (tmpSpace === null) { out.spaces = 0; } else { out.spaces = tmpSpace.length; }
        resetPrefs();
        app.findGrepPreferences.findWhat = this.exclude;
        var findSymbols;
        if (footnote) {
            for (a = 0; a < obj.paragraphs.length; a++) {
                app.findChangeObjectOptions.includeFootnotes = false;
                findSymbols = obj.paragraphs[a].findGrep();
                out.symbols = findSymbols.length;
            }
        } else {
            app.findChangeObjectOptions.includeFootnotes = true;
            findSymbols = obj.findGrep();
            out.symbols = findSymbols.length;
        }
        resetPrefs();
        return out;
    }

    function resetPrefs() {
        app.findGrepPreferences = null;
        app.findGrepPreferences = null;
        app.changeGrepPreferences = null;
    }

    function ProgressbarClass (maxVal, barLabel) {
        var win = new Window("palette", barLabel, [150, 150, 600, 260]);
        this.windowRef = win;
        win.pnl = win.add("panel", [10, 10, 440, 100], "");

        win.pnl.progBarLabel = win.pnl.add("statictext", [20, 15, 380, 30], "Calculating...");
        win.pnl.progBar = win.pnl.add("progressbar", [20, 35, 405, 60], 0, maxVal);

        win.stopButton = win.add("button", [150, 110, 250, 140], "Stop");
        win.stopButton.onClick = function ()  { this.terminated = true; };
        this.terminated = false;

        win.onClose = function() {}

        win.center();
        win.show();

        this.close = function() { win.close(); };
        this.reset = function (newMaxVal, newBarLabel) {
            this.windowRef.pnl.progBar.maxvalue = newMaxVal;
            this.windowRef.pnl.progBar.value = 0;
            this.windowRef.pnl.progBarLabel.text = newBarLabel;
        };
        this.setVal = function (progBarVal) { this.windowRef.pnl.progBar.value = progBarVal; };
        this.increase = function () { this.windowRef.pnl.progBar.value++; };
        this.setLabel = function (St) { this.windowRef.pnl.progBarLabel.text = St; };
        this.setPanelLabel = function (St) { this.windowRef.pnl.text = St; };
    }
}
