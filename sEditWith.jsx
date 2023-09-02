var myDoc = app.activeDocument;

// var myTool = app.toolBoxTools; //// НЕ РАБОТАЕТ
// var myOldTool = myTool;
// myTool.currentTool = UITools.DIRECT_SELECTION_TOOL;
// переключаем на инструмент Direct Selection


//GET SELECTION
var mySel = app.selection[0];
var myFileExt = '';

if (app.selection != 0) {

    switch (app.selection[0].constructor.name) {
        case "Oval":
        case "Polygon":
        case "Rectangle":
            alert(app.selection[0]);
            myFilePath = mySel.allGraphics[0].itemLink.filePath;

            //alert(myFilePath);
            //alert(myFilePath.split('.').reverse()[0]);
            myFileExt = myFilePath.split('.').reverse()[0];
            break;
        case "Image":
            //alert(app.selection[0]);
            //alert(app.selection[0].parent);
            myFilePath = mySel.parent.allGraphics[0].itemLink.filePath;
            //alert(myFilePath);
            //alert(myFilePath.split('.').reverse()[0]);
            myFileExt = myFilePath.split('.').reverse()[0];
            break;
        default:
            alert("Выбранный объект не содержит изображения");
            exit();
    }

    myFileExt = myFileExt.toLowerCase();
    alert(myFileExt);
    switch (myFileExt) {
        case "eps":
        case "ai":
        case "pdf":
          // alert(app.menuActions.itemByName("Adobe Illustrator 2023 27.7 (default)").title);
            //alert(app.menuActions.itemByName("Adobe Illustrator 2023 27.7 (default)").id);
            
            //app.menuActions.itemByName("Adobe Illustartor 2023 27.1").invoke();
            app.menuActions.itemByName("Adobe Illustrator 2023 27.7 (default)").invoke();
            //app.menuActions[132811].invoke();
            break;
        case "jpg":
        case "jpeg":
        case "jpe":
        case "bmp":
        case "png":
        case "tif":
        case "psd":
            //alert(app.menuActions.itemByName("Adobe Photoshop 2023 24.7 (default)").title);
            //alert(app.menuActions.itemByName("Adobe Photoshop 2023 24.7 (default)").id);
            //app.menuActions[132812].invoke();
            //app.menuActions.itemByName("Adobe Photoshop 2023 24.1").invoke();
            app.menuActions.item("$ID/Adobe Photoshop 2023 24.7 (default)").invoke();
            break;
        default:
            alert("Незнакомое расширение файла");
    }
}
else {
    alert('нет выделения');
}



    
/*
var menuItem = app.menus.item("$ID/ParaStylePanelPopup").menuItems.item("$ID/Redefine Style");
if(menuItem.isValid == false)
*/




// app.menuActions.itemByName("Adobe Photoshop 2023 24.1").invoke();

/* скрипт вывода всеъ menuaction

var myActions = app.menuActions;
var myActionsList = Array();
var counter = Number(0);
 
for(var i = 0; i < myActions.length; i++){
    myActionsList.push(String(myActions[i].name));
    myActionsList.push(String(myActions[i].area));
    myActionsList.push(String(myActions[i].id));
}
 
var myDoc = app.activeDocument;
var myTextFrame = myDoc.pages[0].textFrames.add();
myTextFrame.geometricBounds = app.activeDocument.pages[0].bounds;
 
var myMenuActionsTbl = myTextFrame.insertionPoints[0].tables.add();
myMenuActionsTbl.columnCount = 3;
myMenuActionsTbl.bodyRowCount = myActions.length;
myMenuActionsTbl.contents = myActionsList;

*/