// Vitaly Batushev win2utf

function d2h(d) {return d.toString(16);} function h2d(h) {return parseInt(h,16);}  
with (app) {  
    if (documents.length == 0) { exit(); }  
    if (selection.length > 0) { 
        for (s = 0; s < selection.length; s ++) { try { cyr2utf(selection[s]); } catch(e) {} } 
    } else { 
        cyr2utf(activeDocument); 
    } 
} 
function cyr2utf(obj) { 
    app.findGrepPreferences = null; app.changeGrepPreferences = null; 
    app.findGrepPreferences.findWhat = '\\x{00B8}'; app.changeGrepPreferences.changeTo = 'ё'; obj.changeGrep();  
    app.findGrepPreferences.findWhat = '\\x{00A8}'; app.changeGrepPreferences.changeTo = 'Ё'; obj.changeGrep();  
    for (i = 192; i < 256; i++) { app.findGrepPreferences.findWhat = '\\x{00' + d2h(i).toUpperCase() + '}';  app.changeGrepPreferences.changeTo = '\\x{0' + d2h(i + 848).toUpperCase()  + '}'; obj.changeGrep(); } 
    app.findGrepPreferences = null; app.changeGrepPreferences = null;  
}  