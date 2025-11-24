/*
   openWithGuides.jsx
   Открывает связанный файл в Photoshop / Illustrator / InCopy,
   переносит пересекающие направляющие и добавляет контурные
   по границам фрейма. Направляющие строятся в реальных пикселях
   Photoshop с учётом dpi документа.
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

    /* ---------- 2. путь и расширение (с точкой) ---------- */
    var link     = img.itemLink;
    var filePath = link.filePath;
    var name     = File(filePath).name;
    var lastDot  = name.lastIndexOf('.');
    var ext      = (lastDot === -1) ? "" : name.substr(lastDot).toLowerCase();

    /* ---------- 3. выбираем host ---------- */
    var host = "";
    if (".psd .jpg .jpeg .tif .tiff .png".indexOf(ext) > -1)       host = "photoshop";
    else if (".ai .eps .svg".indexOf(ext) > -1)                     host = "illustrator";
    else if (".icml .incx .icap .indd".indexOf(ext) > -1)           host = "incopy";
    else { alert("Неизвестный тип файла: " + ext); return; }

    /* ---------- 4. границы фрейма (points) ---------- */
    var frame  = img.parent;
    var page   = frame.parentPage;
    var gb     = frame.geometricBounds;   // [y1, x1, y2, x2]
    var t = gb[0], l = gb[1], b = gb[2], r = gb[3];

    /* ---------- 5. формируем направляющие для Photoshop (с учётом dpi) ---------- */
    var guidesJS = [
        'var f = new File("' + filePath.replace(/\\/g, '/') + '");',
        'if (!f.exists) throw new Error("File not found");',
        'var doc = app.open(f);',
        'var dpi = doc.resolution;',                       // реальное разрешение
        'function pt2px(pt) { return Math.round(pt * dpi / 72); }', // точки → пиксели
        ''
    ].join('\n');

    /* 5a. пересекающие направляющие InDesign → Photoshop */
    for (var i = 0; i < app.activeDocument.guides.length; i++) {
        var g = app.activeDocument.guides[i];
        if (g.parentPage !== page) continue;

        var posPt = g.location; // points
        if (g.orientation === "Horizontal") {
            if (posPt >= t && posPt <= b) guidesJS += 'doc.guides.add(Direction.HORIZONTAL, pt2px(' + posPt + '));\n';
        } else if (g.orientation === "Vertical") {
            if (posPt >= l && posPt <= r) guidesJS += 'doc.guides.add(Direction.VERTICAL,   pt2px(' + posPt + '));\n';
        }
    }

    /* 5b. контурные направляющие по границам фрейма */
    guidesJS += 'doc.guides.add(Direction.HORIZONTAL, pt2px(' + t + '));\n';
    guidesJS += 'doc.guides.add(Direction.HORIZONTAL, pt2px(' + b + '));\n';
    guidesJS += 'doc.guides.add(Direction.VERTICAL,   pt2px(' + l + '));\n';
    guidesJS += 'doc.guides.add(Direction.VERTICAL,   pt2px(' + r + '));\n';

    /* ---------- 6. Illustrator / InCopy корректировки ---------- */
    if (host === "illustrator") {
        guidesJS = guidesJS.replace(/UnitValue\(/g, '');                // убираем UnitValue
        guidesJS = guidesJS.replace(/pt2px\(/g,        '');             // Illustrator работает в пунктах
        guidesJS = guidesJS.replace(/\)/g,             '');             // убираем лишние скобки
    }
    if (host === "incopy") {
        guidesJS = guidesJS.replace(/doc\.guides\.add/g, 'doc.guides.add')
                           .replace(/Direction\./g,    'GuideDirection.');
    }

    /* ---------- 7. отправляем ---------- */
    var bt = new BridgeTalk();
    bt.target = host;
    bt.body = [
        'var f = new File("' + filePath.replace(/\\/g, '/') + '");',
        'if (!f.exists) throw new Error("File not found");',
        'var doc = app.open(f);',
        host === "incopy" ? 'app.activate(); doc.windows[0].activate();'
                          : 'app.bringToFront(); doc.activate();',
        guidesJS
    ].join('\n');

    bt.onError = function (e) { alert("Ошибка доставки:\n" + e.body); };
    bt.send();
})();