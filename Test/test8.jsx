/*
   openUniversal.jsx
   Открывает связанный файл в Photoshop / Illustrator / InCopy
   + переводит фокус на приложение и документ.
   Корректно обрабатывает точки в имени файла.
   2025-06-25
*/
(function () {
    /* ---------- 1. находим первую подходящую картинку ---------- */
    var img = null;
    for (var i = 0; i < app.selection.length; i++) {
        var item = app.selection[i];
        if (item.constructor.name === "Image") { img = item; break; }
        if (item.allGraphics && item.allGraphics.length > 0) { img = item.allGraphics[0]; break; }
    }
    if (!img || !img.itemLink) { alert("Выделите фрейм/изображение с ПРИВЯЗАННЫМ файлом."); return; }

    /* ---------- 2. надёжное извлечение расширения С точкой ---------- */
    var link     = img.itemLink;
    var filePath = link.filePath;
    var name     = File(filePath).name;
    var lastDot  = name.lastIndexOf('.');
    var ext      = (lastDot === -1) ? "" : name.substr(lastDot).toLowerCase(); // ".jpg" | ".icml" etc.

    /* ---------- 3. выбираем host ---------- */
    var host = "";
    if (".psd .jpg .jpeg .tif .tiff .png".indexOf(ext) > -1)       host = "photoshop";
    else if (".ai .eps .svg".indexOf(ext) > -1)                     host = "illustrator";
    else if (".icml .incx .icap .indd".indexOf(ext) > -1)           host = "incopy";
    else { alert("Неизвестный тип файла: " + ext); return; }

    /* ---------- 4. открываем + фокус ---------- */
    var bt = new BridgeTalk();
    bt.target = host;
    bt.body = [
        'var f = new File("' + filePath.replace(/\\/g, '/') + '");',
        'if (!f.exists) throw new Error("File not found");',
        'var doc = app.open(f);',
        host === "incopy" ? 'app.activate(); doc.windows[0].activate();'
                          : 'app.bringToFront(); doc.activate();'
    ].join('\n');

    bt.onError = function (e) { alert("Ошибка доставки:\n" + e.body); };
    bt.send();
})();