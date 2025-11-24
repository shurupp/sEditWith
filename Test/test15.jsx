/*
  openWithGuides_min.jsx  2025-06-25
  Минимальный рабочий вариант: открывает файл в Photoshop и ставит направляющие.
  Логирует всё в файл и альерт.  ES3-совместимый.
*/
(function () {
  /* 1. находим изображение */
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

  /* 2. путь и расширение */
  var path = img.itemLink.filePath;
  var ext = File(path).name.split(".").pop().toLowerCase();
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

  /* 3. границы фрейма (мм) и страница */
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

  /* 4. перевод мм → пиксели по effective-dpi */
  var dpi = img.effectivePpi[0];
  var mm2px = dpi / 25.4;

  /* 5. направляющие внутри ВИДИМОЙ части файла */
  var guidesCode = [
    '$.writeln("");',
    '$.writeln("========  MINI LOG  ========");',
    '$.writeln("effectivePpi: ' + dpi + '");',
    '$.writeln("mm2px: ' + mm2px + '");',
    'var f = new File("' + path.replace(/\\/g, "/") + '");',
    'if (!f.exists) throw new Error("File not found");',
    "var doc = app.open(f);",
    "var attempts = 0; while (!app.activeDocument && attempts < 50) { $.sleep(100); attempts++; }",
    'if (!app.activeDocument) throw new Error("Document not ready");',
    "",
  ].join("\n");

  /* 5a. контурные направляющие (внутри видимой области) */
  var tPx = 0;
  var bPx = Math.round((b - t) * mm2px);
  var lPx = 0;
  var rPx = Math.round((r - l) * mm2px);

  guidesCode += '$.writeln("border T mm: ' + t + ' → PS px: " + ' + tPx + ");";
  guidesCode += "doc.guides.add(Direction.HORIZONTAL, " + tPx + ");\n";
  guidesCode += '$.writeln("border B mm: ' + b + ' → PS px: " + ' + bPx + ");";
  guidesCode += "doc.guides.add(Direction.HORIZONTAL, " + bPx + ");\n";
  guidesCode += '$.writeln("border L mm: ' + l + ' → PS px: " + ' + lPx + ");";
  guidesCode += "doc.guides.add(Direction.VERTICAL,   " + lPx + ");\n";
  guidesCode += '$.writeln("border R mm: ' + r + ' → PS px: " + ' + rPx + ");";
  guidesCode += "doc.guides.add(Direction.VERTICAL,   " + rPx + ");\n";
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
