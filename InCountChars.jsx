/*
 * @Author: Vitaly Batushev
 * @Date:   2015-09-01 09:25:06
 * @Version:    3.0.1
 * @Last Modified by: Vitaly Batushev
 * @Last Modified time: 2023-08-21 18:02:58
*/

var InCounts = (function() {
    var excludeSymbols = "~-";
    var chars = {spaces: 0, symbols: 0, total: 0, lines: 0}

    var strings = {
        name: 'inCountChars 3.0.1',
        minWords: 'Minimum words in Story',
        viewWithOutSpaces: 'Show numbers characters without spaces',
        viewWithOutSpecial: 'Show numbers characters without special symbols',
        attention: 'Attention!',
        noOpenDocument: 'No open documents!',
        report: 'Report',
        publication: 'Publication',
        lines: 'Lines',
        excludeStory: 'Excluded stories, where characters less than',
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
        credits: ' by Vitaly Batushev\nhttps://github.com/vbatushev/InDesignScripts',
        bars: '===============================',
    }

    var config = {
        exclude: excludeSymbols,
        min: 0,
        spaces: false,
        progress: null
    }

    var run = function() {
        if (app.documents.length < 1) { alert(strings.noOpenDocument, strings.attention); exit(); }
        app.scriptPreferences.enableRedraw = true;
        runDialog();
        config.progress = new ProgressbarClass(app.activeDocument.stories.length, strings.name);
        for (var a = 0,l = app.activeDocument.stories.length; a < l; a++) {
            processStory(app.activeDocument.stories[a])
            config.progress.increase();
        }
        save();

        function processStory(story) {
            if (story.words.length >= config.min || story.tables.length) {
                getChars(story, false);
                for (var a = 0, l = story.tables.length; a < l; a++) {
                    processCells(story.tables[a]);
                }
                processFootnote(story);
                getLines(story);
            }
        }

        function processCells(table) {
            for (var a = 0, l = table.cells.length; a < l; a++) {
                getChars(table.cells[a],false);
            }
        }

        function processFootnote(story) {
            for (var a = 0, l = story.footnotes.length; a < l; a++) {
                getChars(story.footnotes[a], false);
            }
        }

        function getLines(story) {
            for (var a = 0, l = story.textContainers.length; a < l; a++) {
                if (story.textContainers[a].lines !== null) {
                    chars.lines += story.textContainers[a].lines.length;
                }
            }
        }
    }

    var save = function () {
        var aList = Math.round(chars.total/40000);
        if (aList < 1) { aList = strings.lessOneList; };
        var report = strings.bars + '\n' + strings.report + '\n' + strings.bars + '\n\n' +
        strings.publication + ': ' + app.activeDocument.name + '\n';
        if (config.min) {
            report += strings.excludeStory + " " + config.min + "\n";
        }
        report += "\n" + chars.total + ' ' + strings.charsWithSpaces + '.' + '\n';

        if (config.spaces){
            report += (chars.total - chars.spaces) + ' ' + strings.charsWithoutSpaces + '.' + '\n\n';
        }
        if (chars.symbols > 0) {
            report += (chars.total - chars.symbols) + ' ' + strings.charsWithoutSymbols + " " + config.exclude + '.' + '\n\n';
        }
        if (chars.lines > 0) {
            report += strings.lines + ': ' + chars.lines + '.' + '\n\n';
        }

        report = report + strings.authorLists + ': ' + aList + '.' + '\n\n' + strings.bars + '\n' +
        strings.createdWith + ' ' + strings.name + '\n' + strings.credits;
        config.progress.close();
        alert(report, strings.name);
        try {
            var reportFile = new File (app.activeDocument.filePath + "/report.txt");
            reportFile.encoding = 'UTF8';
            reportFile.open("w");
            reportFile.write(report);
            reportFile.close();
            reportFile.execute();
        } catch(e) {}
    }

    var getChars = function(obj, footnote) {
        var content = obj.contents;
        chars.total += obj.characters.length;
        var  re = /\s/g;
        var tmpSpace = content.match(re);
        if (tmpSpace !== null) { chars.spaces += tmpSpace.length; }
        resetGrepPreferences();
        app.findGrepPreferences.findWhat = config.exclude;
        var findSymbols;
        if (footnote) {
            for (var a = 0, l = obj.paragraphs.length; a < l; a++) {
                app.findChangeObjectOptions.includeFootnotes = false;
                findSymbols = obj.paragraphs[a].findGrep();
                chars.symbols += findSymbols.length;
            }
        } else {
            app.findChangeObjectOptions.includeFootnotes = true;
            findSymbols = obj.findGrep();
            chars.symbols += findSymbols.length;
        }
        resetGrepPreferences();
    }

    function resetGrepPreferences() {
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

    function runDialog() {
        var dlg = app.dialogs.add({name:strings.name});
        with (dlg) {
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
        config.exclude = outSmbl.editContents;
        config.spaces = chkWithOutSpace.checkedState;
        config.min = minWords.editValue;
    }

    return {
        run: run,
        config: config
    }
})();
InCounts.run();