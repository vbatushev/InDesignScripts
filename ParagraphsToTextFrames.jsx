/*
* @Author: Vitaly Batushev
* @Date:   2016-08-05 16:09:07
* @Last Modified by:   Vitaly Batushev
* @Last Modified time: 2016-08-08 18:11:57
*/

#target "indesign"

var Paragraphs2TextFrames = (function(){
    /**
     * Конфигурация исполняемого модуля
     * Вот не надо здесь руками править.
     * Для этого есть возможность изменить данный конфиг извне, см. ниже вызов скрипта
     */
    var config = {
        style: "",
        objectStyleName: "",
        objectStyle: null
    };

    /**
     * Основная функция скрипта
     */
    var main = function() {
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

        if (app.documents.length == 0) { return; }
        var searchedParagraphStyle = app.activeDocument.paragraphStyles.itemByName(config.style);
        if (!searchedParagraphStyle.isValid) { return; }
        config.objectStyle = app.activeDocument.objectStyles.itemByName(config.objectStyleName);
        if (!config.objectStyle.isValid) { return; }

        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.appliedParagraphStyle = searchedParagraphStyle;
        app.findGrepPreferences.findWhat = "^(.|\\n)*\\r";
        var founded = app.findGrep(true);
        app.findGrepPreferences = app.changeGrepPreferences = null;

        for (var a = 0, l = founded.length; a < l; a++) {
            processParagraph(founded[a]);
        }
        alert("Обработано " + l + " абзацев стиля " + "\"" + config.style + "\"");
    }

    /**
     * Обработка абзаца
     * @param  {Paragraph} p  Объект абзаца, который будет помещен в текстовый фрейм
     */
    var processParagraph = function(p) {
        var txt = p.texts[0],
            bounds = getBoundsForTextFrame(p),
            parentTextFrame = p.parentTextFrames[0];
        p.insertionPoints.lastItem().contents = "\r";

        var nextParagraph = p;
        var tf = nextParagraph.insertionPoints.lastItem().textFrames.add({
            geometricBounds: bounds,
            appliedObjectStyle: config.objectStyle,
        });
        txt.move(LocationOptions.UNKNOWN, tf.parentStory);
        tf.parentStory.characters.lastItem().remove()
        clearParentParagraph(nextParagraph);

        /**
         * Убираем всякое у абзаца, к которому привязывается создаваемый текстовый фрейм
         * @param  {Paragraph} p  Объект абзаца, который будет помещен в текстовый фрейм
         */
        function clearParentParagraph(p) {
            p.leftIndent = 0;
            p.rightIndent = 0;
            p.firstLineIndent = 0;
            p.lastLineIndent = 0;
        }

        /**
         * Получение границ текстового фрейма
         * @param  {Paragraph} p  Объект абзаца, который будет помещен в текстовый фрейм
         * @return {Array}        Массив границ
         */
        function getBoundsForTextFrame(p) {
            var bounds = [0, 0, 0, 0],
                firstChar = p.characters.firstItem(),
                lastChar = p.characters.lastItem(),
                parentTextFrame = p.parentTextFrames[0];
            var height = firstChar.leading;
            if (height == Leading.AUTO) {
                height = firstChar.pointSize * (firstChar.autoLeading / 100);
            }
            var top = firstChar.baseline - height;

            bounds[1] = parentTextFrame.geometricBounds[1];
            bounds[3] = parentTextFrame.geometricBounds[3];
            bounds[0] = top;
            bounds[2] = lastChar.baseline + config.objectStyle.textFramePreferences.insetSpacing[0] + config.objectStyle.textFramePreferences.insetSpacing[2] + config.objectStyle.strokeWeight;
            return bounds;
        }
    }

    return {
        config: config,
        execute: main
    }
})();

// Указываем стиль искомых абзацев, если такого стиля не будет обнаружено, работа скрипта завершится
Paragraphs2TextFrames.config.style = "Glossary_Text";
// Указываем стиль текстового фрейма для абзацев, если такого стиля не будет обнаружено, работа скрипта завершится
Paragraphs2TextFrames.config.objectStyleName = "ForParagraph";
// Выполняем скрипт через doScript, чтобы можно было вернуть все волшебство одним Ctrl + Z
app.doScript(Paragraphs2TextFrames.execute, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, 'Волшебство');
// Очистка workspace
Paragraphs2TextFrames = null;
delete Paragraphs2TextFrames;
$.gc();$.gc();