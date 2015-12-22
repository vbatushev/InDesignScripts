/*
* @Author: vbatushev
* @Date:   2015-12-19 18:27:30
* @Last Modified by:   vbatushev
* @Last Modified time: 2015-12-19 19:09:38
*/

if (app.documents.length == 0) {
    alert("Нет открытых документов.", "Внимание!", true);
    exit();
}

var dlg = new DialogTemplateSelect("Выбор шаблона страницы");

var start = true;
for (var a = 0; a < app.activeDocument.pages.length; a++) {
    var page = app.activeDocument.pages[a];
    if (page.appliedMaster == dlg.master) {
        if (start) {
            page.select(SelectionOptions.REPLACE_WITH);
            start = false;
        } else {
            page.select(SelectionOptions.ADD_TO);
        }
    }
}


function DialogTemplateSelect(title) {
    var res = "dialog {\
        properties: {resizeable: false, maximizeButton: false}, \
        dlg: Group { orientation: 'column', alignChildren: ['fill', 'fill'], spacing:10, \
            panel: Panel { orientation: 'row', alignChildren:['left', 'center'], bounds: [90, 90, 370, 160], margins:[10, 20, 10, 20], text: 'Шаблоны страниц в документе', \
                tpl: DropDownList {bounds:[70, 70, 320, 40]}, \
            }, \
        }, \
        btns: Group { orientation: 'row', alignChildren: ['fill', 'fill'], spacing:10, \
            ok: Button { name: 'ok', text:'Выделить страницы', alignment:'center' }, \
            cancel: Button { text:'Закрыть', name: 'cancel', alignment:'center' }, \
        } \
    }";

    var win = new Window(res);
    win.text = title;
    win.defaultElement = win.btns.ok;

    var tpls = {};
    for (var a = 0; a < app.activeDocument.masterSpreads.length; a++) {
        win.dlg.panel.tpl.add("item", app.activeDocument.masterSpreads[a].name);
        tpls[app.activeDocument.masterSpreads[a].name] = app.activeDocument.masterSpreads[a];
    }
    win.dlg.panel.tpl.selection = 0;

    win.btns.cancel.onClick = function() {
        win.close();
        exit();
    };

    win.center();
    var result = win.show();
    if (result !== 1) {
        exit();
    } else {
        this.master = tpls[win.dlg.panel.tpl.selection];
    }

}