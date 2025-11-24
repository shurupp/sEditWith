/*
   openLinkedImage_universal.jsx
   Универсальный opener для InDesign → Photoshop / Illustrator
   Версия 2025-06-25
   Работает с любым выделением: рамка (чёрная стрелка), само изображение (белая стрелка),
   несколько объектов (берёт первый подходящий).
   Направляющие НЕ переносятся.
*/
(function () {
    /* ---------- 1. Находим первую пригодную картинку в выделении ---------- */
    var img = null;
    for (var i = 0; i < app.selection.length; i++) {
        var item = app.selection[i];

        // случай 1: выделено само изображение
        if (item.constructor.name === "Image") {
            img = item;
            break;
        }
        // случай 2: выделена рамка — ищем первую картинку внутри
        if (item.allGraphics && item.allGraphics.length > 0) {
            img = item.allGraphics[0];
            break;
        }
    }

    if (!img || !img.itemLink) {
        alert("Выделите рамку или изображение с ПРИВЯЗАННЫМ файлом.");
        return;
    }

    /* ---------- 2. Определяем Photoshop / Illustrator ---------- */
    var link     = img.itemLink;
    var filePath = link.filePath;
    var ext      = File(filePath).name.toLowerCase().slice(-4);

    var targetApp = "";
    if (".psd .jpg jpeg .tif tiff .png".indexOf(ext) > -1) {
        targetApp = "photoshop";
    } else if (".ai .eps .svg".indexOf(ext) > -1) {
        targetApp = "illustrator";
    } else {
        alert("Неизвестный тип файла: " + ext);
        return;
    }

    /* ---------- 3. Открываем файл ---------- */
    var bt = new BridgeTalk();
    bt.target = targetApp;
    bt.body   = 'app.open(new File("' + filePath.replace(/\\/g, "/") + '"));';
    bt.onError = function (e) { alert("Не удалось открыть файл:\n" + e.body); };
    bt.send();
})();