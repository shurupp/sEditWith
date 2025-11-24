/*
  openWithGuides_final.jsx  2025-06-25
  ES3-совместимый код: открывает связанный файл в Photoshop / Illustrator / InCopy,
  переносит пересекающие направляющие и добавляет контурные по границам фрейма.
  Координаты переводятся из мм InDesign → пиксели Photoshop (с учётом dpi)
  и корректно переворачиваются для Illustrator.
*/
(function () {
  /* 1. находим изображение в выделении */
  var img = null;
  for (var i = 0; i < app.selection.length; i++) {
    var it = app.selection[i];
    if (it.constructor.name === "Image") {
      img = it;
      break;
    }
    if (it.allGraphics && it.allGraphics.length) {
      img = it.allGraphics[0];
      break;
    }
  }
  if (!img || !img.itemLink) {
    alert("Выделите фрейм/изображение с ПРИВЯЗАННЫМ файлом.");
    return;
  }

  /* 2. путь и расширение (без точки) */
  var link = img.itemLink;
  var path = link.filePath;
  var ext = File(path).name.split(".").pop().toLowerCase();

  /* 3. выбираем host */
  var host = "";
  if (".psd .jpg .jpeg .tif .tiff .png".indexOf(ext) > -1) host = "photoshop";
  else if (".ai .eps .svg".indexOf(ext) > -1) host = "illustrator";
  else if (".icml .incx .icap .indd".indexOf(ext) > -1) host = "incopy";
  else {
    alert("Неизвестный тип файла: " + ext);
    return;
  }

  /* 4. границы фрейма (мм) и страница */
  var frame = img.parent;
  var page = frame.parentPage;
  if (!page) {
    alert("Фрейм не находится на странице.");
    return;
  }
  var gb = frame.geometricBounds; // [y1, x1, y2, x2] в мм
  var t = gb[0],
    l = gb[1],
    b = gb[2],
    r = gb[3];

  /* ---------- 5. видимая область изображения + матрица преобразования ---------- */
  var vb = img.visibleBounds; // видимая часть в единицах документа
  var mtx = img.resolve(
    AnchorPoint.TOP_LEFT_ANCHOR,
    CoordinateSpaces.PASTEBOARD_COORDINATES
  )[0]; // матрица
  var scaleX = mtx[0][0],
    scaleY = mtx[1][1]; // масштаб по X и Y
  var offsetX = mtx[0][2],
    offsetY = mtx[1][2]; // сдвиг в пунктах
  var dpiEff = img.effectivePpi[0]; // effective dpi файла
  var docUnits = app.activeDocument.viewPreferences.horizontalMeasurementUnits;
  var pxPerUnit =
    {
      pixels: 1,
      points: 1,
      picas: 12,
      inches: 72,
      inchesDecimal: 72,
      millimeters: 72 / 25.4,
      centimeters: 72 / 2.54,
    }[docUnits] || 1;
  var ptPerUnit = pxPerUnit; // 1 px = 1 pt при 72 dpi

  var guidesCode = [
    '$.writeln("");',
    '$.writeln("========  GUIDES LOG  ========");',
    '$.writeln("Units ID: ' + docUnits + '");',
    '$.writeln("effectivePpi: ' + dpiEff + '");',
    '$.writeln("scaleX: ' + scaleX + ", scaleY: " + scaleY + '");',
    '$.writeln("offsetX(pt): ' + offsetX + ", offsetY(pt): " + offsetY + '");',
    'var f = new File("' + path.replace(/\\/g, "/") + '");',
    'if (!f.exists) throw new Error("File not found");',
    "var doc = app.open(f);",
    "var attempts = 0; while (!app.activeDocument && attempts < 50) { $.sleep(100); attempts++; }",
    'if (!app.activeDocument) throw new Error("Document not ready");',
    "var dpi = doc.resolution;", // реальное dpi файла
    "function id2px(val) { return Math.round(val * " +
      pxPerUnit +
      " * dpi / 72); }",
    "function id2pt(val) { return val * " + ptPerUnit + "; }",
    "",
  ].join("\n");

  /* 5a. пересекающиеся направляющие */
  for (var i = 0; i < app.activeDocument.guides.length; i++) {
    var g = app.activeDocument.guides[i];
    if (g.parentPage !== page) continue;
    var pos = g.location; // в единицах документа
    if (g.orientation === "Horizontal") {
      if (pos >= t && pos <= b) {
        /* переводим в пиксели относительно ВИДИМОЙ области */
        var relPt = (pos - vb[0]) * scaleY + offsetY; // pt относительно видимой части
        if (host === "photoshop") {
          var px = Math.round((relPt * dpi) / 72);
          guidesCode +=
            '$.writeln("ID H guide: ' +
            pos +
            " mm → rel pt: " +
            relPt +
            ' → PS px: " + ' +
            px +
            ");";
          guidesCode += "doc.guides.add(Direction.HORIZONTAL, " + px + ");\n";
        } else if (host === "illustrator") {
          var yIll = page.bounds[2] - pos; // переворот Y
          guidesCode +=
            '$.writeln("ID H guide: ' + pos + " mm → AI pt: " + yIll + '");';
          guidesCode +=
            "doc.guides.add(GuideDirection.HORIZONTAL, " + yIll + ");\n";
        }
      }
    } else if (g.orientation === "Vertical") {
      if (pos >= l && pos <= r) {
        var relPt = (pos - vb[1]) * scaleX + offsetX; // pt относительно видимой части
        if (host === "photoshop") {
          var px = Math.round((relPt * dpi) / 72);
          guidesCode +=
            '$.writeln("ID V guide: ' +
            pos +
            " mm → rel pt: " +
            relPt +
            ' → PS px: " + ' +
            px +
            ");";
          guidesCode += "doc.guides.add(Direction.VERTICAL, " + px + ");\n";
        } else if (host === "illustrator") {
          guidesCode +=
            '$.writeln("ID V guide: ' + pos + " mm → AI pt: " + pos + '");';
          guidesCode +=
            "doc.guides.add(GuideDirection.VERTICAL, " + pos + ");\n";
        }
      }
    }
  }

    /* 5b. контурные (внутри BridgeTalk, где уже есть 'var dpi') */
    guidesCode += '$.writeln("border T mm: ' + t + '");';
    guidesCode += 'var tPx = Math.round(((' + ((vb[0] - t) * scaleY + offsetY)) + ') * dpi / 72);';
    guidesCode += '$.writeln("→ PS px: " + tPx);';
    guidesCode += 'doc.guides.add(Direction.HORIZONTAL, tPx);\n';

    guidesCode += '$.writeln("border B mm: ' + b + '");';
    guidesCode += 'var bPx = Math.round(((' + ((vb[2] - t) * scaleY + offsetY)) + ') * dpi / 72);';
    guidesCode += '$.writeln("→ PS px: " + bPx);';
    guidesCode += 'doc.guides.add(Direction.HORIZONTAL, bPx);\n';

    guidesCode += '$.writeln("border L mm: ' + l + '");';
    guidesCode += 'var lPx = Math.round(((' + ((vb[1] - l) * scaleX + offsetX)) + ') * dpi / 72);';
    guidesCode += '$.writeln("→ PS px: " + lPx);';
    guidesCode += 'doc.guides.add(Direction.VERTICAL, lPx);\n';

    guidesCode += '$.writeln("border R mm: ' + r + '");';
    guidesCode += 'var rPx = Math.round(((' + ((vb[3] - l) * scaleX + offsetX)) + ') * dpi / 72);';
    guidesCode += '$.writeln("→ PS px: " + rPx);';
    guidesCode += 'doc.guides.add(Direction.VERTICAL, rPx);\n';

  guidesCode += '$.writeln("========  END LOG  ========");';

  /* 6. отправляем одним сообщением */
  var bt = new BridgeTalk();
  bt.target = host;
  bt.body = guidesCode;
  bt.onError = function (e) {
    alert("Ошибка доставки:\n" + e.body);
  };
  bt.send();
})();
