/**
 * openLinkedImage_clean.jsx
 * Упрощённый вариант первого скрипта: только открывает связанный файл
 * в Photoshop (растр) или Illustrator (вектор). Направляющие НЕ переносятся.
 * Выделяйте рамку или сам рисунок и запускайте.
 */
(function () {
    /* ---------- 1. Получаем картинку из выделения ---------- */
    if (app.selection.length !== 1) {
        alert("Выделите ОДИН фрейм с изображением.");
        return;
    }

    var container = app.selection[0];
    var img = null;

    if (container.constructor.name === "Image") {
        img = container;                                    // кликнули по картинке
    } else if (container.allGraphics && container.allGraphics.length > 0) {
        img = container.allGraphics[0];                     // берём первую картинку в рамке
    }

    if (!img || !img.itemLink) {
        alert("В выделенной рамке нет ПРИВЯЗАННОГО изображения.");
        return;
    }

    /* ---------- 2. Определяем целевое приложение ---------- */
    var link     = img.itemLink;
    var filePath = link.filePath;




var ext = File(filePath).name.toLowerCase().slice(-4);

var targetApp = "";
if (".psd .jpg jpeg .tif tiff .png".indexOf(ext) > -1) {
    targetApp = "photoshop";
} else if (".ai .eps .svg".indexOf(ext) > -1) {
    targetApp = "illustrator";
} else {
    alert("Неизвестный тип файла.");
    return;
}






    /* ---------- 3. Открываем файл ---------- */
    var bt = new BridgeTalk();
    bt.target = targetApp;
    bt.body   = 'app.open(new File("' + filePath.replace(/\\/g,"/") + '"));';
    bt.onError = function (e) { alert("Не удалось открыть файл:\n" + e.body); };
    bt.send();
})();