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

  /* ---------- 5. мм рамки → пиксели файла + лог ---------- */
  var dpi = img.effectivePpi[0];
  var mm2px = dpi / 25.4;

  /* переводим 2 точки изображения → страница → pasteboard */
  var pImg00 = [0, 0]; // левый-верхний угол изображения
  var pImg11 = [1, 1]; // 1 мм вниз и вправо

  /* 1. изображение → страница (через parent страницы) */
  var pPage00 = img.resolve(pImg00, CoordinateSpaces.INNER_COORDINATES);
  var pPage11 = img.resolve(pImg11, CoordinateSpaces.INNER_COORDINATES);

  /* 2. страница → pasteboard (через parent страницы) */
  var pBoard00 = page.resolve(pPage00, CoordinateSpaces.PARENT_COORDINATES);
  var pBoard11 = page.resolve(pPage11, CoordinateSpaces.PARENT_COORDINATES);

  var scaleX = pBoard11[0] - pBoard00[0];
  var scaleY = pBoard11[1] - pBoard00[1];
  var offsetX = pBoard00[0];
  var offsetY = pBoard00[1];

  /* лог-файл на рабочем столе */
  var logFile = new File(
    Folder.desktop + "/guides_log_" + new Date().getTime() + ".txt"
  );
  logFile.open("w");
  logFile.writeln("========  GUIDES LOG  ========");
  logFile.writeln("effectivePpi: " + dpi);
  logFile.writeln("mm2px: " + mm2px);
  logFile.writeln("scaleX: " + scaleX + ", scaleY: " + scaleY);
  logFile.writeln("offsetX(pt): " + offsetX + ", offsetY(pt): " + offsetY);

  var alerts = [];

  /* 5a. пересекающиеся направляющие → пиксели от ЛЕВОГО-ВЕРХНЕГО УГЛА ИЗОБРАЖЕНИЯ */
  for (var i = 0; i < app.activeDocument.guides.length; i++) {
    var g = app.activeDocument.guides[i];
    if (g.parentPage !== page) continue;
    var posMm = g.location;
    if (g.orientation === "Horizontal") {
      if (posMm >= t && posMm <= b) {
        var relPt = (posMm - t) * scaleY + offsetY;
        var px = Math.round(relPt * mm2px);
        var msg =
          "ID H guide mm: " + posMm + " → rel pt: " + relPt + " → PS px: " + px;
        logFile.writeln(msg);
        alerts.push(msg);
        guidesCode += "doc.guides.add(Direction.HORIZONTAL, " + px + ");\n";
      }
    } else if (g.orientation === "Vertical") {
      if (posMm >= l && posMm <= r) {
        var relPt = (posMm - l) * scaleX + offsetX;
        var px = Math.round(relPt * mm2px);
        var msg =
          "ID V guide mm: " + posMm + " → rel pt: " + relPt + " → PS px: " + px;
        logFile.writeln(msg);
        alerts.push(msg);
        guidesCode += "doc.guides.add(Direction.VERTICAL, " + px + ");\n";
      }
    }
  }

  /* 5b. контурные */
  var tPx = Math.round(offsetY);
  var bPx = Math.round((b - t) * scaleY * mm2px + offsetY);
  var lPx = Math.round(offsetX);
  var rPx = Math.round((r - l) * scaleX * mm2px + offsetX);

  var tMsg = "border T mm: " + t + " → abs px: " + tPx;
  var bMsg = "border B mm: " + b + " → abs px: " + bPx;
  var lMsg = "border L mm: " + l + " → abs px: " + lPx;
  var rMsg = "border R mm: " + r + " → abs px: " + rPx;

  logFile.writeln(tMsg);
  logFile.writeln(bMsg);
  logFile.writeln(lMsg);
  logFile.writeln(rMsg);
  alerts.push(tMsg, bMsg, lMsg, rMsg);

  logFile.writeln("========  END LOG  ========");
  logFile.close();

  /* выводим кратко в алерт + путь к файлу */
  alerts.push("", "Лог сохранён: " + logFile.fsName);
  alert(alerts.join("\n"));

  /* 6. отправляем одним сообщением */
  var bt = new BridgeTalk();
  bt.target = host;
  bt.body = guidesCode;
  bt.onError = function (e) {
    alert("Ошибка доставки:\n" + e.body);
  };
  bt.send();
})();
