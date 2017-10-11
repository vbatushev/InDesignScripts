/*
 * @Author: Vitaly Batushev
 * @Date: 2017-10-11 14:23:14
 * @Last Modified by: Vitaly Batushev
 * @Last Modified time: 2017-10-11 15:19:54
 *
 * Порой возникают ситуации, когда в InDesign появляются богомерзкие ссылки на изображения,
 * которые не входят в компоновку, а являются структурным элементом XML.
 * В сведения о странице написано: XML structure element, not in layout
 * (Структурный элемент XML, в компоновку не входит).
 * Исправить весь этот ужасъ и смрадъ призван данный скрипт.
 *
 * Основано на обсуждении https://forums.adobe.com/thread/781135
 * Основной код Loic.Aigon(https://forums.adobe.com/people/Loic.Aigon)
 */
#target "indesign"
var XMLinksCleaner = (function () {
    var main = function() {
        var question = "Вы собираетесь удалить ссылки, которые не входят в компоновку, являясь структурными элементами XML?";
        question += "\nЕсли Вы не знаете или не понимаете, что Вы собираетесь делать, закройте этот скрипт.";
        question += "\nПосле чего прикройте глаза и трижды произнесите \"Ссылка-ссылка рассосись, будет скоро забишись\".";
        question += "\nЕсли, открыв глаза, Вы все же обнаружите ссылки, у которых вместо номера страницы стоят знаки < >,";
        question += " запустите скрипт заново и ответьте \"Да\". После чего снова можете произнести заклятие.";

        if (!app.documents.length) return;
        if (!confirm(question, true, "XMLinksCleaner")) return;
        var doc = app.properties.activeDocument, root, story, items, n, rect;
            root  = doc.xmlElements[0],
            story = root.parentStory,
            items = story.pageItems,
            n     = items.length;
        while (n--) {
            rect = items[n].getElements()[0];
            !rect.properties.geometricBounds && rect.remove();
        }
        doc.save();
        alert("Обработка и исправление закончено.\nПубликация сохранена.", "XMLinksCleaner");
    }
    return { run: main }
})();
app.doScript("XMLinksCleaner.run();", ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "XMLinksCleaner");