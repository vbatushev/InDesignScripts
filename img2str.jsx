/*
* @Author: Vitaly Batushev
* @E-Mail: vitaly@batushev.info
* @Date:   2016-03-19 22:01:09
* @Last Modified by:   Vitaly Batushev
* @Last Modified time: 2016-03-19 22:08:54
* @ScriptName  img2str
* @Version     1.0
* @Description Создание текстового представления изображения для использования его внутри скрипта
*/

var img2str = (function(){
    var main = function() {
        var imgs = File.openDialog ("Выберите файл", "*.png, *.jpg, *.jpeg", true);
        if (imgs === null) {
            exit();
        }

        for (var a = 0, l = imgs.length; a < l; a++) {
            createResourceFile(imgs[a]);
        }
        alert("Текстовые файлы из выбранных изображений созданы.");

        function createResourceFile(img) {
            img.open("r");
            img.encoding = "BINARY";
            var imgstr = img.read();
            img.close();

            var outfile = new File(img.absoluteURI + ".txt");
            outfile.open("w");
            outfile.write(imgstr.toSource().replace(/^\(new String\("/gm,"").replace(/"\)\)$/gm, ""));
            outfile.close();
        }
    }

    return {
        run:main
    }
})();
img2str.run();
