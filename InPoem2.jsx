/*
 * @Author: Vitaly Batushev
 * @Date: 2017-02-10 15:22:28
 * @Last Modified by: Vitaly Batushev
 * @Last Modified time: 2025-02-27 19:14:44
 * @Name: InPoem
 * @Version: 2.1.0
 * @Email: vitaly@batushev.info
 * @Description:
 * Скрипт выравнивания строк стихотворения по самой длинной строке в нем
 * Обходит все текстовые фреймы, относящиеся к материалу (Story) выделенного текстового фрейма,
 * где ищет абзацы со стилями, указанными в параметрах head_styles (для заголовков стихотворений)
 * и text_styles (для основного текста стихотворений), которые и попадают под выравнивание
*/
#target indesign

/**
 * Конфигурация стилей
 * head_style — стиль заголовков стихов
 * text_style — стиль основного текста стихов
 */
var styles = {
    head_styles: ["Заголовок стихотворения"],
    text_styles: ["Стихи"]
}

var InPoem = (function () {
    var config = { head_styles: [], text_styles: [] }
    var Index = { start: 0, end: 0 }

    /**
     * Основная функция объекта InPoem
     */
    var main = function () {
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
        testPublication();
        processTextFrames(app.selection[0].parentStory);
        alert("Стихи обработаны.", "InPoem 2.1.1");
    }

    /**
     * Проверка публикации на возможность выполнения скрипта
     * При невозможности скрипт сообщает о причине отказа
     */
    var testPublication = function () {
        if (app.documents.length < 1) { alert("Нет открытых публикаций!\nОткрой публикацию и попробуйте снова.", "Ошибка", true); exit(); }
        if (app.selection.length < 1) { alert("Ничего не выделено!\nВыделите текстовый фрейм и попробуйте снова.", "Ошибка", true); exit(); }
        if (app.selection[0].constructor.name != "TextFrame") {
            alert("Выделенный объект не является текстовым фреймов!\nВыделите текстовый фрейм и попробуйте снова.", "Ошибка", true); exit();
        }
        for (var i = 0; i < config.head_styles.length; i++) {
            if (!app.activeDocument.paragraphStyles.itemByName(config.head_styles[i]).isValid) {
                alert("Стиль " + config.head_styles[i] + " не найден!", "Ошибка", true);
                exit();
            }
        }
        for (var i = 0; i < config.text_styles.length; i++) {
            if (!app.activeDocument.paragraphStyles.itemByName(config.text_styles[i]).isValid) {
                alert("Стиль " + config.text_styles[i] + " не найден!", "Ошибка", true);
                exit();
            }
        }
    }

    /**
     * Обработка текстового фрейма, относящихся к parentStory
     * @param  {Story}  story   Объект Story, текстовые фреймы которого необходимо обработать
     */
    var processTextFrames = function (story) {
        var textContainersLength = story.textContainers.length;
        var progress = new ProgressbarClass(textContainersLength, "Обрабатываются текстовые фреймы", "InPoem 2.0", false);
        for (var a = 0, l = story.textContainers.length; a < l; a++) {
            progress.setLabel("Обрабатывается текстовый фрейм " + a + " из " + textContainersLength);
            var textFrame = story.textContainers[a];
            if (textFrame.characters.length == 0) { continue; }
            var re = new RegExp("([^\u000D]+)|($)", "gim");
            var heads = seekRegexp(textFrame, re);

            Index.start = heads[0][0]; Index.end = 0;
            for (var i = 0, j = heads.length; i < j; i++) {
                if (heads[i][1] < 0) { continue; }
                if (!textFrame.characters[heads[i][1]].isValid) { continue; }
                var typeStyle = getTypeOfStyle(textFrame.characters[heads[i][1]].appliedParagraphStyle.name);
                switch (typeStyle) {
                    case "head":
                        if (Index.end > 0) {
                            setIndents(textFrame);
                        }
                        Index.start = heads[i][1] + 1;
                        break;
                    case "text":
                        Index.end = heads[i][1];
                        if (i == heads.length - 1 && Index.end > 0) {
                            setIndents(textFrame);
                        }
                        break;
                    default:
                        if (Index.end > 0) setIndents(textFrame);
                        break;
                }
            }
            progress.increase();
        }
        progress.close();
    }

    /**
     * Установка отступов
     * в текстовом фрейме
     * @param {TextFrame}   tf      Обрабатываемых текстовый фрейм
     * @param {Integer}     start   Индекс первого символа текста
     * @param {Integer}     end     Индекс последнего символа текста
     */
    function setIndents(tf) {
        var text = tf.characters.itemByRange(Index.start, Index.end);
        var indent = calcIndent(text, tf);
        indent = Math.round(indent * 10) / 10;
        for (var a = 0, l = text.lines.length; a < l; a++) {
            var line = text.lines[a];
            if (!line.isValid) continue;
            if (line.hasOwnProperty("appliedParagraphStyle")) {
                var type;
                if (!line.appliedParagraphStyle.name) {
                    type = getTypeOfStyle(line.appliedParagraphStyle[0].name);
                } else {
                    type = getTypeOfStyle(line.appliedParagraphStyle.name);
                }

                if (type == 'head' || type == 'text') {
                    line.leftIndent = indent;
                    line.rightIndent = indent;
                }
            }
        }
        Index.start = 0; Index.end = 0;
    }

    /**
     * Вычисление величины отступа
     * для данного текста (Text)
     * @param   {Text}        text   Обрабатываемый текст
     * @param   {TextFrame}   tf     Текстовый фрейм
     * @returns {Integer}            Величина отступа в пт
     */
    function calcIndent(text, tf) {
        var values = [];
        for (var a = 0, l = text.lines.length; a < l; a++) {
            var paragraph = text.lines[a];
            var type = getTypeOfStyle(paragraph.appliedParagraphStyle[0].name);
            if (type == 'head' || type == 'text') {
                var startIP = paragraph.insertionPoints.item(0);
                var endIP = paragraph.insertionPoints.item(paragraph.insertionPoints.length - 2);
                values.push(endIP.horizontalOffset - startIP.horizontalOffset);
            }
        }
        if (values.length == 0) { return 0; }
        values.sort(function (a, b) { return a - b; });
        width = tf.geometricBounds[3] - tf.geometricBounds[1];
        return (width - values[values.length - 1]) / 2;
    }

    /**
     * Определение границ по регулярному выражению
     *
     * @param   {Text}    text   Объект Text
     * @param   {RegExp}  re     Регулярное выражение
     * @return  {Array}   Массив из найденных границ начала и конца
     */
    function seekRegexp(tf, re) {
        var i = 0, result = [], text = tf.contents, found = text.match(re);
        if (!found) { return result; }
        for (var a = 0, l = found.length; a < l; a++) {
            var start = text.indexOf(found[a], i);
            var end = start + (found[a].length - 1);
            var type = getTypeOfStyle(tf.characters[start].appliedParagraphStyle.name);
            if (end > start && (type == 'head' || type == 'text')) {
                result.push([start, end]);
                i = end;
            }
        }
        return result;
    }

    /**
     * Получение типа стиля абзаца
     *
     * @param   {String} styleName    Имя стиля
     * @return  {String}            "head", "text" или "other"
     */
    function getTypeOfStyle(styleName) {
        for (var a = 0, l = config.head_styles.length; a < l; a++) {
            if (config.head_styles[a] == styleName) {
                return "head";
            }
        }
        for (var a = 0, l = config.text_styles.length; a < l; a++) {
            if (config.text_styles[a] == styleName) {
                return "text";
            }
        }
        return "other";
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

    function ProgressbarClass(maxValue, barLabel, panelTitle, viewInfo) {
        var panelTitle = panelTitle || Locales.ProgressBar.panelTitle;
        var viewInfo = viewInfo || false;
        var Info = "";

        var win = new Window("palette", "", undefined, { closeButton: true, maximizeButton: false, minimizeButton: false, resizeable: false, borderless: false });
        this.windowRef = win;

        if (viewInfo) {
            win.pnl = win.add("panel", [10, 10, 440, 210], panelTitle);
            win.pnl.progBarLabel = win.pnl.add("statictext", [20, 15, 405, 30], barLabel);
            win.pnl.progBar = win.pnl.add("progressbar", [20, 35, 405, 60], 0, maxValue);
            win.pnl.log = win.pnl.add("panel", [20, 65, 405, 175], "");
            win.pnl.log.info = win.pnl.log.add("statictext", [5, 5, 385, 150], Info, { multiline: true, scrolling: false });
            win.pnl.log.info.visible = true;
        } else {
            win.pnl = win.add("panel", [10, 10, 440, 100], panelTitle);
            win.pnl.progBarLabel = win.pnl.add("statictext", [20, 15, 405, 30], barLabel);
            win.pnl.progBar = win.pnl.add("progressbar", [20, 35, 405, 60], 0, maxValue);
        }

        win.cancelBtn = win.add("button", [0, 0, 0, 0], "Отменить");
        win.cancelBtn.visible = true;
        win.cancelElement = win.cancelBtn;
        this.terminated = false;

        win.onCancel = function () {
            this.isDone = true;
            return true;
        }

        win.center();
        win.show();

        this.close = function () {
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

    return {
        config: config,
        run: main
    }
})();

InPoem.config.head_styles = styles.head_styles;
InPoem.config.text_styles = styles.text_styles;
app.doScript("InPoem.run();", ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "InPoem 2");
// InPoem.run();