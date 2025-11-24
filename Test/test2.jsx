/*
  OpenLinkedImage_withGuides.jsx
  Автор: @klimovsky
  2025-06-25
  InDesign → Photoshop / Illustrator + перенос направляющих
  Работает при выделенной рамке (Rectangle, GraphicFrame и т.д.)
*/

(function () {
    if (app.selection.length !== 1) {
        alert("Выделите ОДИН фрейм с изображением.");
        return;
    }

    /* 1. Получаем картинку внутри выделенной рамки -------------------- */
    var container = app.selection[0];
    var img = null;

    if (container.constructor.name === "Image") {
        img = container;                                    // пользователь кликнул по картинке
    } else if (container.allGraphics && container.allGraphics.length > 0) {
        img = container.allGraphics[0];                     // берём первую картинку в рамке
    }

    if (!img || !img.itemLink) {
        alert("В выделенной рамке нет ПРИВЯЗАННОГО изображения.");
        return;
    }

    /* 2. Определяем целевое приложение ------------------------------- */
    var link      = img.itemLink;
    var filePath  = link.filePath;
    var ext       = File(filePath).name.toLowerCase().substr(-4);

    var targetApp = "";
    if ([".psd",".jpg","jpeg",".tif","tiff",".png"].some(function(e){return ext.indexOf(e)>-1;}))
        targetApp = "photoshop";
    else if ([".ai",".eps",".svg"].some(function(e){return ext.indexOf(e)>-1;}))
        targetApp = "illustrator";
    else { alert("Неизвестный тип файла."); return; }

    /* 3. Границы рамки (в координатах страницы) ---------------------- */
    var doc   = app.activeDocument;
    var page  = container.parentPage;
    var gb    = container.geometricBounds;
    var t = gb[0], l = gb[1], b = gb[2], r = gb[3];

    /* 4. Собираем пересекающие направляющие -------------------------- */
    var guidesJS = "";
    for (var i = 0; i < doc.guides.length; i++) {
        var g = doc.guides[i];
        if (g.parentPage !== page) continue;

        var pos = g.location;
        if (g.orientation === GuideOrientation.HORIZONTAL) {
            if (pos >= t && pos <= b)
                guidesJS += 'doc.guides.add(Direction.HORIZONTAL, UnitValue(' + pos + '));\n';
        } else {
            if (pos >= l && pos <= r)
                guidesJS += 'doc.guides.add(Direction.VERTICAL,   UnitValue(' + pos + '));\n';
        }
    }

    /* 5. BridgeTalk: открываем файл → создаём направляющие ---------- */
    var btOpen = new BridgeTalk();
    btOpen.target = targetApp;
    btOpen.body   = 'app.open(new File("' + filePath.replace(/\\/g,"/") + '"));';
    btOpen.onError = function (e) { alert("Не удалось открыть файл:\n" + e.body); };

    btOpen.onResult = function () {
        if (!guidesJS) return;               // направляющих нет
        var btGuide = new BridgeTalk();
        btGuide.target = targetApp;
        btGuide.body   = 'var doc = app.activeDocument;\n' + guidesJS;
        btGuide.onError = function (e) { alert("Ошибка при создании направляющих:\n" + e.body); };
        btGuide.send();
    };

    btOpen.send();
})();