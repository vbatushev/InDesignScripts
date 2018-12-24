/*
 * @Author: Vitaly Batushev
 * @Date: 2017-02-16 17:52:35
 * @Last Modified by: Vitaly Batushev
 * @Last Modified time: 2018-12-24 14:06:00
 */
#target indesign

var ImageConvertor = (function() {
    var scriptVersion = "1.7";
    /**
     * Объект с функциями, передаваемые через BridgeTalk
     */
    var exportfunc = {
        ai: ' \
        function convertEPS(path, exportPath) { \
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS; \
        \
            var doc = app.open(File(path), DocumentColorSpace.CMYK);\
            var tf = doc.textFrames;\
            for (t = tf.length - 1; t >= 0; t--) {\
                var text = tf[t];\
                text.createOutline();\
            }\
            var saveOptions = new IllustratorSaveOptions(); \
            saveOptions.compatibility = Compatibility.ILLUSTRATOR16; \
            saveOptions.compressed = true; \
            saveOptions.embedICCProfile = false; \
            saveOptions.pdfCompatible = true; \
            saveOptions.fontSubsetThreshold = 100; \
            doc.selectObjectsOnActiveArtboard(); \
            doc.fitArtboardToSelectedArt (0); \
            doc.saveAs(File(exportPath)); \
            doc.close(SaveOptions.DONOTSAVECHANGES); \
        }',
        eps: ' \
        function convertEPS(path, exportPath) { \
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS; \
        \
            var doc = app.open(File(path), DocumentColorSpace.CMYK);\
            var tf = doc.textFrames;\
            for (t = tf.length - 1; t >= 0; t--) {\
                var text = tf[t];\
                text.createOutline();\
            }\
            var saveOptions = new EPSSaveOptions(); \
            saveOptions.compatibility = Compatibility.ILLUSTRATOR17; \
            saveOptions.compatibleGradientPrinting = false; \
            saveOptions.cmykPostScript = true; \
            saveOptions.embedAllFonts = false; \
            saveOptions.embedLinkFiles = false; \
            saveOptions.includeDocumentThumbnails = false; \
            saveOptions.overprint = PDFOverprint.PRESERVEPDFOVERPRINT; \
            saveOptions.postScript = EPSPostScriptLevelEnum.LEVEL2; \
            saveOptions.preview = EPSPreview.NONE; \
            doc.selectObjectsOnActiveArtboard(); \
            doc.fitArtboardToSelectedArt (0); \
            doc.saveAs(File(exportPath), saveOptions); \
            doc.close(SaveOptions.DONOTSAVECHANGES); \
        }',
        tif: ' \
        function convertEPS(path, exportPath) { \
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS; \
        \
            var doc = app.open(File(path), DocumentColorSpace.CMYK);\
            var exp = new ExportOptionsTIFF(); \
            exp.imageColorSpace = ImageColorSpace.CMYK; \
            exp.embedICCProfile = false; \
            exp.lZWCompression = false; \
            exp.saveMultipleArtboards = false; \
            exp.resolution = 300; \
            doc.selectObjectsOnActiveArtboard(); \
            doc.fitArtboardToSelectedArt (0); \
            doc.exportFile(File(exportPath), ExportType.TIFF, exp); \
            doc.close(SaveOptions.DONOTSAVECHANGES); \
        }',
        psd: ' \
        function convertEPS(path, exportPath) { \
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS; \
        \
            var doc = app.open(File(path), DocumentColorSpace.CMYK);\
            var exp = new ExportOptionsPhotoshop(); \
            exp.antiAliasing = true; \
            exp.imageColorSpace = ImageColorSpace.CMYK; \
            exp.editableText = false; \
            exp.embedICCProfile = false; \
            exp.maximumEditability = false; \
            exp.resolution = 300; \
            exp.saveMultipleArtboards = false; \
            exp.writeLayers = false; \
            doc.selectObjectsOnActiveArtboard(); \
            doc.fitArtboardToSelectedArt (0); \
            doc.exportFile(File(exportPath), ExportType.PHOTOSHOP, exp); \
            doc.close(SaveOptions.DONOTSAVECHANGES); \
        }'
    }

    // TODO: Сделать создание списка несконвертированных файлов

    /**
     * EPSlinks     — массив объектов с параметрами source (полное имя исходного файла в формате URI)
     *                и target (полное имя генерируемого файла в формате URI).
     *                Необходим для передачи через BridgeTalk.
     * Progress     — окно, показывающее процесс выполнения
     * Format       — тип экспортируемого файла, заодно его расширение и имя параметра в объекте exportfunc
     * relinkMode   — режим только перелинковки файлов (если перелинкуемые файлы уже присутствуют)
     * epssount     — количество найденных EPS и WMF
     */
    var EPSlinks = {}, Progress = null, Format = "ai", relinkMode = false, epscount = 0, report = "";

    /**
     * Основная функция программы
     */
    var main = function() {
        dialogFormat();
        Progress = new ProgressbarClass(1, "Поиск EPS и WMF…", "Конвертор EPS и WMF " + scriptVersion, false);
        getAllEPSLinks();
        if (relinkMode) {
            updateAll();
        } else {
            if (!!epscount) processLinks();
        }
        Progress.close();
    }

    /**
     * Обработка ссылок из массива getAllEPSLinks
     */
    function processLinks() {
        Progress.reset(epscount, "Обработка EPS и WMF…");
        var a = 1;
        for (l in EPSlinks) {
            // Вызываем Adobe Illustrator
            var targetApp = BridgeTalk.getSpecifier( "illustrator");
            if (targetApp) {
                var bt = new BridgeTalk;
                bt.target = targetApp;
                bt.onResult = function(returnBtObj) {
                    // По окончании работы скрипта в Illustrator
                    // переключается в InDesign и перелинковываем
                    BridgeTalk.bringToFront('indesign');
                    if (a == epscount) {
                        updateAll();
                    }
                    Progress.increase();
                }
                bt.body = exportfunc[Format] + "\nconvertEPS(\"" + l + "\", \"" +  EPSlinks[l].target.absoluteURI + "\");";
                bt.send(60);
            }
            a++;
        }
    }

    /**
     * Полное обновление ссылок в публикации InDesign
     */
    function updateAll() {
        Progress.reset(epscount, "Перелинковка…");
        for (l in EPSlinks) {
            if (EPSlinks[l].target.exists) {
                EPSlinks[l].source.relink(EPSlinks[l].target);
            } else {
                report += File(l).name + "\n";
            }
            Progress.increase();
        }

        Progress.reset(app.activeDocument.links.length, "Обновление линков…");
        for (var a = 0, l = app.activeDocument.links.length; a < l; a++) {
            var link = app.activeDocument.links[a];
            if (link.status == LinkStatus.NORMAL || link.status == LinkStatus.LINK_OUT_OF_DATE) {
                link.update();
            }
            Progress.increase();
        }
        Progress.close();
        if (report != "") {
            report = "Не обработаны ссылки на:\n" + report;
        } else {
            report = "Перелинковка закончена."
        }
        alert(report);
    }

    /**
     * Получение всех ссылок на файлы EPS
     *
     */
    function getAllEPSLinks() {
        Progress.reset(app.activeDocument.allGraphics.length, "Поиск EPS и WMF…");
        for (var a = 0, l = app.activeDocument.allGraphics.length; a < l; a++) {
            var g = app.activeDocument.allGraphics[a];
            if (g instanceof EPS || g instanceof WMF) {
                var link = g.itemLink;
                if (link.status == LinkStatus.NORMAL || link.status == LinkStatus.LINK_OUT_OF_DATE) {
                    putInEPSlinks(link);
                }
            }
            Progress.increase();
        }

        /**
         * Помещение объекта о ссылке в EPSLinks
         * @param {Link} link Линк в InDesign
         */
        function putInEPSlinks(link) {
            var source = File(link.filePath).absoluteURI;
            if (!!EPSlinks[source]) return;
            if (source.substr (-11) == "_relink.eps") return;
            var targetFile;
            if (Format == "eps") {
                targetFile = File(getNameWithoutExtension(link.filePath) + "_relink."+ Format);
            } else {
                targetFile = File(getNameWithoutExtension(link.filePath) + "." + Format);
            }
            EPSlinks[source] = {source: link, target: targetFile};
            epscount++;
        }
        /**
         * Получение имени файла без расширения
         * @param   {String} name     Полное имя файла
         * @returns {String}          Имя файла без расширения
         */
        function getNameWithoutExtension(name) {
            return name.split(".").slice(0,-1).join(".");
        }
    }

    /**
     * Класс индикатора выполнения
     * @param {Number} maxValue             Максимальное значение индикатора
     * @param {String} barLabel             Текст над индикатором
     * @param {String} panelTitle           Заголовок панели
     * @param {String} Info                 Текст под индикатором
     *
     * @method close()                      Закрывает окно индикатора выполнения
     * @method reset(maxValue, barLabel)    Новые значения для максимума и текста над индикатором
     * @method setVal(barValue)             Устанавливает конкретное значение индикатора
     * @method increase()                   Увеличивает значение индикатора на единицу
     * @method setLabel(str)                Устанавливает новое значение текста над индикатором
     * @method setPanelLabel(str)           Устанавливает новое значение заголовка панели
     * @method setInfo(str)                 Устанавливает новое значение текста под индикатором
     */

    function ProgressbarClass (maxValue, barLabel, panelTitle, viewInfo) {
        var panelTitle = panelTitle || Locales.ProgressBar.panelTitle;
        var viewInfo = viewInfo || false;
        var Info = "";

        var win = new Window("palette", "", undefined, { closeButton:true, maximizeButton:false, minimizeButton:false, resizeable:false, borderless: false});
        this.windowRef = win;

        if (viewInfo) {
            win.pnl = win.add("panel", [10, 10, 440, 210], panelTitle);
            win.pnl.progBarLabel = win.pnl.add("statictext", [20, 15, 405, 30], barLabel);
            win.pnl.progBar = win.pnl.add("progressbar", [20, 35, 405, 60], 0, maxValue);
            win.pnl.log = win.pnl.add("panel", [20, 65, 405, 175], "");
            win.pnl.log.info = win.pnl.log.add("statictext", [5, 5, 385, 150], Info, {multiline: true, scrolling: false});
            win.pnl.log.info.visible = true;
        } else {
            win.pnl = win.add("panel", [10, 10, 440, 100], panelTitle);
            win.pnl.progBarLabel = win.pnl.add("statictext", [20, 15, 405, 30], barLabel);
            win.pnl.progBar = win.pnl.add("progressbar", [20, 35, 405, 60], 0, maxValue);
        }

        win.cancelBtn  = win.add("button", [0, 0, 0, 0], "Отменить");
        win.cancelBtn.visible = true;
        win.cancelElement = win.cancelBtn ;
        this.terminated = false;

        win.onCancel = function() {
            this.isDone = true;
            return true;
        }

        win.center();
        win.show();

        this.close = function() {
            win.close();
        };
        this.reset = function (maxValue, barLabel) {
            this.windowRef.pnl.progBar.maxvalue = maxValue;
            this.windowRef.pnl.progBar.value = 0;
            this.windowRef.pnl.progBarLabel.text = barLabel;
        };
        this.setVal = function (barValue) { this.windowRef.pnl.progBar.value = barValue; };
        this.increase = function () { this.windowRef.pnl.progBar.value++; };
        this.setLabel = function (str) { this.windowRef.pnl.progBarLabel.text = str; };
        this.setPanelLabel = function (str) { this.windowRef.pnl.text = str; };
        this.setInfo = function (str) {
            this.windowRef.pnl.log.info.text = str;
        };
    }

    /**
     * Вывод диалогового окна с выбором формата
     * генерируемых файлов для перелинковки
     */
    function dialogFormat() {
        var res = "dialog { \
            properties: {resizeable: false, maximizeButton: false, text: 'Конвертор EPS' }, alignChildren: ['fill', 'fill'], \
            topblock: Panel { orientation: 'column', spacing:10, margins:[10, 20, 10, 20],  alignChildren: ['fill', 'fill'], text: 'Выберите формат:', \
                eps: RadioButton { text:'Illustrator EPS', value: true }, \
                ai: RadioButton { text:'Adobe Illustrator', value: false }, \
                psd: RadioButton { text:'Adobe Photoshop', value: false }, \
                tif: RadioButton { text:'TIFF', value: false }, \
                onlyrelink: Checkbox { text:'Только перелинковка', value: false }, \
            }, \
            btns: Group { orientation: 'row', alignChildren: ['fill', 'fill'], spacing:10, \
                ok: Button { name: 'ok', text:'Конвертировать', alignment:'center', enabled: true }, \
                cancel: Button { text:'Отменить', name: 'cancel', alignment:'center' }, \
            } \
        }";
        var win = new Window(res);
        win.text = "Конвертор EPS/WMF " +  + scriptVersion;
        win.defaultElement = win.btns.ok;
        win.btns.cancel.onClick = function() {
            win.close();
            exit();
        };
        win.center();
        var result = win.show();
        if (result !== 1) {
            exit();
        } else {
            Format = "eps";
            if (win.topblock.ai.value) Format = "ai";
            if (win.topblock.tif.value) Format = "tif";
            if (win.topblock.psd.value) Format = "psd";
            relinkMode = win.topblock.onlyrelink.value;
        }
    }

    return {
        run: main
    }
})();
ImageConvertor.run();