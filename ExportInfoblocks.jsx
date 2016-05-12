/*
* @Author: Vitaly Batushev
* @Date:   2016-05-08 21:22:53
* @Last Modified by:   Vitaly Batushev
* @Last Modified time: 2016-05-09 10:27:18
*/

/**
 * Экспорт информационных блоков
 * @param  config.styles     {Array}   Массив объектов, описывающих колонки
 *                                     head — заголовок колонки,
 *                                     style — название стиля, абзац с которым будет помещаться в колонку
 * @param config.delimiter   {String}  Разделитель колонок, по умолчанию — ;
 * @param config.encoding    {String}  Кодировка файла, по умолчанию CP1251 (MS Excel любит его)
 */
var ExportInfoblock = (function(){
    var config = {
        styles: [
            {head: "Первый", style: "Заголовок"},
            {head: "Второй", style: "Вводка"},
            {head: "Третий", style: "Normal Буквица"}
        ],
        delimiter: ";",
        encoding: "CP1251"
    };
    // rows — массив рядов будущей таблицы, row — ряд, который будет записываться в файл рапорта
    var rows = [], row = "";
    var main = function() {
        // Если открытых документов нет, прекращаем работу
        if (app.documents.length === 0) { exit(); }

        //Обходим все объекты Story в Document
        for (var a = 0, l = app.activeDocument.stories.length; a < l; a++) {
            processStory(app.activeDocument.stories[a]);
        }

        // Если массив рядов заполнен, записываем файл и сообщаем пользователю, в какой
        // Если массив пуст, сообщаем об этом
        if (rows.length > 0) {
            var fileName = writeReport();
            alert("Отчет записан в файл " + fileName);
        } else {
            alert("Отчет не составлен по причине отсутствия текстов");
        }
    }

    /**
     * Обработка объекта Story
     * @param  {Story} story Обрабатываемый объект
     */
    var processStory = function(story) {
        var temprow = [], pos = 0;
        for (var a = 0, l = story.paragraphs.length; a < l; a++) {
            var current = story.paragraphs[a],
                styleName = current.appliedParagraphStyle.name,
                contents = replaceSpecCharacters(current.contents);
            // Ищем место в ряду
            var place = findPlace(styleName);
            if (place == 0) {
                if (temprow[0] !== undefined) {
                    rows.push(temprow);
                    temprow = [];
                }
                temprow[0] = contents;
                pos = current;
            } else {
                if (temprow[0] !== undefined) {
                    if (temprow[place] === undefined) {
                        temprow[place] = contents;
                    } else {
                        temprow[place] += contents;
                    }
                }
            }
        }
        if (temprow[0] !== undefined) {
             rows.push(temprow);
        }

        /**
         * Поиск места в ряду
         * @param  {String} str  Имя стиля
         * @return {Number}      Позиция в ряду
         */
        function findPlace(str) {
            for (var a = 0, l = config.styles.length; a < l; a++) {
                if (str == config.styles[a].style) {
                    return a;
                }
            }
            // Если стиль не найден в массиве, считаем,
            // что он относится к последнему в массиве стилей
            return config.styles.length - 1;
        }

        /**
         * Замена специальных символов в строке
         * @param  {String}   str   Исходная строка
         * @return {String}         Строка с замененными символами
         */
        function replaceSpecCharacters (str) {
            return str.replace(/\r$/,"").replace(/\n$/,"").replace(/(\n|\r)/," ");
        }
    }

    /**
     * Запись файла рапорта
     * @return {String}     Имя файла рапорта
     */
    var writeReport = function() {
        var fileReport = File(app.activeDocument.filePath + "/" + app.activeDocument.name + "-report.csv");
        fileReport.encoding = config.encoding; fileReport.open("w");
        for (var a = 0, l = config.styles.length; a < l; a++) {
            row += "\"" + config.styles[a].head + "\"" + config.delimiter;
        }
        fileReport.writeln(row);

        for (var a = 0, l = rows.length; a < l; a++) {
            row = "";
            for (var i = 0, j = rows[a].length; i < j; i++) {
                var content = rows[a][i];
                if (content === undefined) {
                    content = "";
                }
                row += "\"" + content + "\"" + config.delimiter;
            }
            fileReport.writeln(row);
        }
        fileReport.close();
        return fileReport.name;
    }

    return {
        execute: main,
        config: config
    }
})();

// Пример переопределения массива стилей
ExportInfoblock.config.styles = [
    {head: "Заголовок статьи", style: "Заголовок"},
    {head: "Вводная часть", style: "Вводка"},
    {head: "Первый абзац", style: "Normal Буквица"},
    {head: "Основной текст", style: "Normal"}
];
// Пример переопределения кодировки
// ExportInfoblock.config.encoding = "UTF8";

// Пример переопределения разделителя
// ExportInfoblock.config.delimiter = ",";

ExportInfoblock.execute();
