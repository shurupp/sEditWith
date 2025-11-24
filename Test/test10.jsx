/*

*/
(function () {
  /* 1. получаем картинку */
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
    alert("Нет связанного изображения");
    return;
  }

  /* 2. рамка и страница */
  var frame = img.parent; // ← сначала рамка
  var page = frame.parentPage; // ← потом страница
  if (!page) {
    alert("Фрейм не на странице");
    return;
  }

  /* ---------- 1. выделение ---------- */
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
    alert("Нет связанного изображения");
    return;
  }

  /* ---------- 2. путь и расширение (с точкой) ---------- */
  var link = img.itemLink;
  var filePath = link.filePath;
  var path = img.itemLink.filePath;
  var name = File(filePath).name;
  var lastDot = name.lastIndexOf(".");
  var ext = lastDot === -1 ? "" : name.substr(lastDot).toLowerCase();

  $.writeln(ext);

  /* ---------- 3. выбираем host ---------- */
  var host = "";
  if (".psd .jpg .jpeg .tif .tiff .png".indexOf(ext) > -1) host = "photoshop";
  else if (".ai .eps .svg".indexOf(ext) > -1) host = "illustrator";
  else if (".icml .incx .icap .indd".indexOf(ext) > -1) host = "incopy";
  else {
    alert("Неизвестный тип файла: " + ext);
    return;
  }
  var frame = img.parent;
  var page = frame.parentPage;
  var gb = frame.geometricBounds; // [y1, x1, y2, x2]
  var t = gb[0],
    l = gb[1],
    b = gb[2],
    r = gb[3];
  /* ---------- 5. переводим мм InDesign → пиксели Photoshop ---------- */
  var ml2pt = 2.83464566929134; // 1 мм = 2,83464566929134 pt
  var guidesCode = [
    'var f = new File("' + path.replace(/\\/g, "/") + '");',
    'if (!f.exists) throw new Error("File not found");',
    "var doc = app.open(f);",
    "var attempts = 0; while (!app.activeDocument && attempts < 50) { $.sleep(100); attempts++; }",
    'if (!app.activeDocument) throw new Error("Document not ready");',
    "var dpi = doc.resolution;",
    "function mm2px(mm) { var pt = mm * " +
      ml2pt +
      "; return Math.round(pt * dpi / 72); }",
    "",
  ].join("\n");

  /* 5a. пересекающиеся направляющие */
  for (var i = 0; i < app.activeDocument.guides.length; i++) {
    var g = app.activeDocument.guides[i];
    if (g.parentPage !== page) continue;
    var posMm = g.location; // миллиметры
    if (g.orientation === "Horizontal") {
      if (posMm >= t && posMm <= b) {
        guidesCode += '$.writeln("PS H guide px = " + mm2px(' + posMm + "));";
        guidesCode +=
          "doc.guides.add(Direction.HORIZONTAL, mm2px(" + posMm + "));\n";
      }
    } else if (g.orientation === "Vertical") {
      if (posMm >= l && posMm <= r) {
        guidesCode += '$.writeln("PS V guide px = " + mm2px(' + posMm + "));";
        guidesCode +=
          "doc.guides.add(Direction.VERTICAL,   mm2px(" + posMm + "));\n";
      }
    }
  }

  /* 5b. контурные */
  guidesCode += '$.writeln("border T px = " + mm2px(' + t + "));";
  guidesCode += '$.writeln("border B px = " + mm2px(' + b + "));";
  guidesCode += "doc.guides.add(Direction.HORIZONTAL, mm2px(" + t + "));";
  guidesCode += "doc.guides.add(Direction.HORIZONTAL, mm2px(" + b + "));";
  guidesCode += "doc.guides.add(Direction.VERTICAL,   mm2px(" + l + "));";
  guidesCode += "doc.guides.add(Direction.VERTICAL,   mm2px(" + r + "));";

  /* ---------- 6. отправляем ---------- */
  var bt = new BridgeTalk();
  bt.target = host;
  bt.body = guidesCode;
  bt.onError = function (e) {
    alert("Ошибка:\n" + e.body);
  };
  bt.send();
})();
