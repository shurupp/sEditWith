// InDesign Script to open linked image in Photoshop or Illustrator and create guides where they intersect the frame

// Main function to execute the script
function main() {
    // Check if any object is selected
    if (app.selection.length == 0) {
        alert("Please select a frame.");
        return;
    }

    // Get the selected frame
    var frame = app.selection[0];

    // Check if the selected object is an image frame
    if (frame.constructor.name != "Image") {
        alert("Please select an Image frame.");
        return;
    }

    // Check if the frame has a linked file
    if (frame.itemLink == null) {
        alert("The selected frame does not contain a linked file.");
        return;
    }

    // Get the linked file path
    var link = frame.itemLink;
    var page = frame.parent;

// получаем путь и имя файла
var filePath = link.filePath;
var fileName = File(filePath).name.toLowerCase();   // только имя файла

// определяем тип
var appToOpen;
if (
    fileName.indexOf(".psd")  > -1 ||
    fileName.indexOf(".jpg")  > -1 ||
    fileName.indexOf(".jpeg") > -1 ||
    fileName.indexOf(".tif")  > -1 ||
    fileName.indexOf(".tiff") > -1 ||
    fileName.indexOf(".png")  > -1
) {
    appToOpen = "photoshop";
} else if (
    fileName.indexOf(".ai")  > -1 ||
    fileName.indexOf(".eps") > -1 ||
    fileName.indexOf(".svg") > -1
) {
    appToOpen = "illustrator";
} else {
    alert("Неизвестный тип файла. Выберите растровое или векторное изображение.");
    return;
}

    // Get geometric bounds of the frame [top, left, bottom, right]
    var frameGB = frame.geometricBounds;
    var frameTop = frameGB[0];
    var frameLeft = frameGB[1];
    var frameBottom = frameGB[2];
    var frameRight = frameGB[3];

// --- исправленный участок ---
var doc   = app.activeDocument;
var myPageIndex = page.documentOffset;          // индекс страницы в документе
var allGuides = doc.guides;                     // все направляющие документа
var intersectingGuides = [];

for (var i = 0; i < allGuides.length; i++) {
    var g = allGuides[i];

    // пропускаем направляющие на других страницах
    if (g.pageIndex !== myPageIndex) continue;

    var loc = g.location;
    if (g.orientation === GuideOrientation.HORIZONTAL) {
        if (loc >= frameTop && loc <= frameBottom) {
            intersectingGuides.push({type:"horizontal", position:loc});
        }
    } else { // вертикальная
        if (loc >= frameLeft && loc <= frameRight) {
            intersectingGuides.push({type:"vertical",   position:loc});
        }
    }
}
// --------------------------------

    // Prepare coordinates to send to Photoshop or Illustrator
    var coordinates = [];
    for (var i = 0; i < intersectingGuides.length; i++) {
        var guide = intersectingGuides[i];
        coordinates.push({
            type: guide.type,
            position: guide.position
        });
    }

    // Create a BridgeTalk message to send to the target application
    var bt = new BridgeTalk();
    bt.target = appToOpen;

    // Construct the script to be executed in the target application
    var scriptBody = 'var doc = app.activeDocument;\n';
    scriptBody += 'if (doc) {\n';
    for (var i = 0; i < coordinates.length; i++) {
        var guide = coordinates[i];
        if (appToOpen == "photoshop") {
            if (guide.type == "horizontal") {
                scriptBody += '    doc.guides.add(Direction.HORIZONTAL, UnitValue(' + guide.position + '));\n';
            } else if (guide.type == "vertical") {
                scriptBody += '    doc.guides.add(Direction.VERTICAL, UnitValue(' + guide.position + '));\n';
            }
        } else if (appToOpen == "illustrator") {
            if (guide.type == "horizontal") {
                scriptBody += '    doc.guides.add(GuideDirection.HORIZONTAL, ' + guide.position + ');\n';
            } else if (guide.type == "vertical") {
                scriptBody += '    doc.guides.add(GuideDirection.VERTICAL, ' + guide.position + ');\n';
            }
        }
    }
    scriptBody += '}';

    bt.body = scriptBody;
    bt.onError = function(inBT) {
        alert("Error communicating with " + appToOpen + ': ' + inBT.body);
    };
    bt.onResult = function(res) {
        alert("Guides created in " + appToOpen + ".");
    };

    // Open the linked file in the target application
    try {
        var openBt = new BridgeTalk();
        openBt.target = appToOpen;
        openBt.body = 'app.open(new File("' + filePath + '"));';
        openBt.onError = function(inBT) {
            alert("Error opening file in " + appToOpen + ": " + inBT.body);
        };
        openBt.onResult = function(res) {
            // Send the guide creation script after the file is opened
            bt.send();
        };
        openBt.send();
    } catch (e) {
        alert("Error opening the linked file: " + e.message);
        return;
    }
}

// Execute the main function
main();