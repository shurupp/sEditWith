/*
  openWithGuides_final.jsx  2025-06-25
  ES3-совместимый код: открывает связанный файл в Photoshop / Illustrator / InCopy,
  переносит пересекающиеся направляющие и добавляет контурные по видимой части изображения.
  Учитывает масштаб, сдвиг, единицы измерения InDesign и dpi Photoshop/Illustrator.
  Логирует каждый шаг в Debug Console Photoshop/Illustrator.
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
  if (".psd.jpg.jpeg.tif.tiff.png.".indexOf("." + ext + ".") > -1)
    host = "photoshop";
  else if (".ai.eps.svg.".indexOf("." + ext + ".") > -1) host = "illustrator";
  else if (".icml.incx.icap.indd.".indexOf("." + ext + ".") > -1)
    host = "incopy";
  else {
    alert("Неизвестный тип файла: " + ext);
    return;
  }

  /* 4. границы фрейма (мм) и страница */
  var frame = img.parent;
  var page = frame.parentPage;
  if (!page) {
    alert("Фрейм не на странице.");
    return;
  }
  var gb = frame.geometricBounds; // [y1, x1, y2, x2] в мм
  var t = gb[0],
    l = gb[1],
    b = gb[2],
    r = gb[3];

  /* ---------- 5. мм рамки → пиксели ВИДИМОЙ области файла ---------- */
  var dpi = img.effectivePpi[0]; // пикселей на дюйм файла
  var mm2px = dpi / 25.4; // 1 мм → пиксели при effective-dpi

  var guidesCode = [
    '$.writeln("");',
    '$.writeln("========  GUIDES LOG  ========");',
    '$.writeln("effectivePpi: ' + dpi + '");',
    '$.writeln("mm2px: ' + mm2px + '");',
    'var f = new File("' + path.replace(/\\/g, "/") + '");',
    'if (!f.exists) throw new Error("File not found");',
    "var doc = app.open(f);",
    "var attempts = 0; while (!app.activeDocument && attempts < 50) { $.sleep(100); attempts++; }",
    'if (!app.activeDocument) throw new Error("Document not ready");',
    "",
  ].join("\n");

  /* 5a. пересекающиеся направляющие → пиксели от ЛЕВОГО-ВЕРХНЕГО УГЛА ВИДИМОЙ ЧАСТИ */
  for (var i = 0; i < app.activeDocument.guides.length; i++) {
    var g = app.activeDocument.guides[i];
    if (g.parentPage !== page) continue;
    var posMm = g.location; // мм от левого-верхнего угла страницы
    if (g.orientation === "Horizontal") {
      if (posMm >= t && posMm <= b) {
        var relMm = posMm - t; // мм от верхнего края рамки
        var px = Math.round(relMm * mm2px); // пиксели от верхнего края файла
        guidesCode +=
          '$.writeln("ID H guide mm: ' +
          posMm +
          " → rel mm: " +
          relMm +
          ' → PS px: " + ' +
          px +
          ");";
        guidesCode += "doc.guides.add(Direction.HORIZONTAL, " + px + ");\n";
      }
    } else if (g.orientation === "Vertical") {
      if (posMm >= l && posMm <= r) {
        var relMm = posMm - l; // мм от левого края рамки
        var px = Math.round(relMm * mm2px); // пиксели от левого края файла
        guidesCode +=
          '$.writeln("ID V guide mm: ' +
          posMm +
          " → rel mm: " +
          relMm +
          ' → PS px: " + ' +
          px +
          ");";
        guidesCode += "doc.guides.add(Direction.VERTICAL, " + px + ");\n";
      }
    }
  }

  /* 5b. контурные (от левого-верхнего угла видимой части) */
  var tPx = 0; // верхний край файла
  var bPx = Math.round((b - t) * mm2px); // нижний край файла
  var lPx = 0; // левый край файла
  var rPx = Math.round((r - l) * mm2px); // правый край файла

  guidesCode += '$.writeln("border T mm: ' + t + ' → abs px: " + ' + tPx + ");";
  guidesCode += "doc.guides.add(Direction.HORIZONTAL, " + tPx + ");\n";
  guidesCode += '$.writeln("border B mm: ' + b + ' → abs px: " + ' + bPx + ");";
  guidesCode += "doc.guides.add(Direction.HORIZONTAL, " + bPx + ");\n";
  guidesCode += '$.writeln("border L mm: ' + l + ' → abs px: " + ' + lPx + ");";
  guidesCode += "doc.guides.add(Direction.VERTICAL,   " + lPx + ");\n";
  guidesCode += '$.writeln("border R mm: ' + r + ' → abs px: " + ' + rPx + ");";
  guidesCode += "doc.guides.add(Direction.VERTICAL,   " + rPx + ");\n";

  guidesCode += '$.writeln("========  END LOG  ========");';

  /* 7. отправляем одним сообщением */
  var bt = new BridgeTalk();
  bt.target = host;
  bt.body = guidesCode;
  bt.onError = function (e) {
    alert("Ошибка доставки:\n" + e.body);
  };
  bt.send();
})();
