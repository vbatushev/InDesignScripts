/*
* @Author: Vitaly Batushev
* @E-Mail: vitaly@batushev.info
* @Date:   2015-12-07 02:43:24
 * @Last Modified by: Vitaly Batushev
 * @Last Modified time: 2017-06-13 13:19:27
* @ScriptName  win2utf
* @Version     1.0
* @Description Скрипт для исправления неуникодных символов, которые появляются при использовании "корявых" шрифтов в импортируемой в InDesign публикации Adobe PageMaker.
*/

var win2utf = (function(){
    var main = function() {
        if (app.documents.length == 0) { exit(); }
        if (app.selection.length > 0) {
            for (var a = 0, l = selection.length; a < l; a++) {
                try {
                    cyr2utf(selection[a]);
                } catch(e) {}
            }
        } else {
            cyr2utf(app.activeDocument);
        }
    }

    var cyr2utf = function(obj) {
        app.findGrepPreferences = null;
        app.changeGrepPreferences = null;

        app.findGrepPreferences.findWhat = "\\x{00B8}";
        app.changeGrepPreferences.changeTo = 'ё';
        obj.changeGrep();

        app.findGrepPreferences.findWhat = "\\x{00A8}";
        app.changeGrepPreferences.changeTo = "Ё";
        obj.changeGrep();

        for (var i = 192, j = 256; i < j; i++) {
            app.findGrepPreferences.findWhat = "\\x{00" + d2h(i).toUpperCase() + "}";
            app.changeGrepPreferences.changeTo = "\\x{0" + d2h(i + 848).toUpperCase()  + "}";
            obj.changeGrep();
        }

        app.findGrepPreferences = null;
        app.changeGrepPreferences = null;
    }

    var d2h = function(d) { return d.toString(16); }
    var h2d = function(h) { return parseInt(h,16); }

    return {
        run: main
    }
})();
win2utf.run();