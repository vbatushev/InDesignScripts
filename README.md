### XMLinksCleaner

Порой возникают ситуации, когда в InDesign появляются богомерзкие ссылки на изображения, которые не входят в компоновку, а являются структурным элементом XML.

В сведениях о странице написано: XML structure element, not in layout (Структурный элемент XML, в компоновку не входит).

Исправить весь этот ужасъ и смрадъ призван данный скрипт.

### InPoem2

Скрипт выравнивания строк стихотворения по самой длинной строке в нем

Обходит все текстовые фреймы, относящиеся к материалу (Story) выделенного текстового фрейма, где ищет абзацы со стилями, указанными в параметрах head_styles (для заголовков стихотворений) и text_styles (для основного текста стихотворений), которые и попадают под выравнивание

### ParagraphsToTextFrames

Скрипт ищет в тексте абзацы указанного пользователем стиля и оборачивает их в текстовый фрейм, который заменяет данный абзац в текстовом потоке.

### ExportInfoblocks

Экспорт текста публикации в формат CSV по информационным блокам согласно стилям, определяемым в массиве ExportInfoblock.config.styles. Исходное задание: http://forum.rudtp.ru/threads/indesign-to-csv.65159/

### InBackup

Создание резервных копий документа в папке указываемой пользователем.

Для использования необходимо (пере)назначить горячую клавишу сохранения документа на этот скрипт.

### InCountChars

Статистика о количестве знаков в публикации.

Можно исключать из подсчета определенные символы, указав их в списке исключаемых символов.

Список составляется в формате GREP. Можно использовать GREP-символы Adobe InDesign.

При перечислении символы разделяются символом |.

Например, нижеследующая настройка исключает из подсчета неразрывный дефис:
> var excludeSymbols = "~-";

### RecoveryINDD

Скрипт для "лечения" или уменьшения размера публикации Adobe InDesign.

Принцип работы в пересборке публикации на основе генерируемого из нее файла IDML.

### selectPagesByMaster

Выделение страниц выбранного шаблона.

### win2utf

Скрипт для исправления неуникодных символов, которые появляются при использовании "корявых" шрифтов в импортируемой в InDesign публикации Adobe PageMaker.

### img2str

Создание текстового представления изображения для использования его внутри скрипта

### ImageConvertor.jsx

Конвертация векторных изображений (EPS и WMF) в EPS/AI/TIFF/PSD с помощью Adobe Illustrator