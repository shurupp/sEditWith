var myDoc = app.activeDocument;

//GET SELECTION
var mySel = app.selection[0];
var myFileExt = '';

if (app.selection != 0) {
    alert(app.selection[0].constructor.name);

    switch (app.selection[0].constructor.name) {
        case "Oval":
        case "Polygon":
        case "Rectangle":
            myFilePath = mySel.allGraphics[0].itemLink.filePath;
            myFileExt = myFilePath.split('.').reverse()[0];
            break;
        case "Image":
        case "PDF":
        case "EPS":
            myFilePath = mySel.parent.allGraphics[0].itemLink.filePath;
            //mySel.parent.select();
            myFileExt = myFilePath.split('.').reverse()[0];
            break;
        default:
            alert("Выбранный объект не содержит изображения");
            exit();
    }

    myFileExt = myFileExt.toLowerCase();
    //alert(myFileExt);

    var phNameVer = "";
    var illNameVer = "";
    var phIDVer ;
    var illIDVer ;

    var a = app.menuActions.everyItem().getElements();

    for (var i = 0; i < a.length; i++) {
        if (a[i].name.search("Adobe Photoshop") != -1) {
            alert(a[i].area + ' : ' + a[i].name, a[i].id);
            phNameVer = a[i].name;
        } 
        if (a[i].name.search("Adobe Illustrator") != -1) {
            alert(a[i].area + ' : ' + a[i].name, a[i].id);
            illNameVer = a[i].name;
            illIDVer = a[i].id;
        }
    }

    switch (myFileExt) {
        case "eps":
        case "ai":
        case "pdf":
            //app.menuActions.itemByName(illNameVer).invoke();
            alert (app.menuActions.itemByName(illNameVer), 'menu');
            alert (app.menuActions.itemByName(illNameVer).name, 'name');

            alert (app.menuActions.itemByID(illIDVer).id, 'id');
            app.menuActions.itemByID(illIDVer).invoke();
            break;
        case "jpg":
        case "jpeg":
        case "jpe":
        case "bmp":
        case "png":
        case "tif":
        case "psd":
        case "heic":
            alert (app.menuActions.itemByName(phNameVer), 'menu');
            alert (app.menuActions.itemByName(phNameVer).name, 'name');

            alert (app.menuActions.itemByID(phIDVer).id, 'id');
            app.menuActions.itemByName(phNameVer).invoke();
            break;
        default:
            alert("Незнакомое расширение файла");
    }
} else {
    alert('Выделите фрейм с изображением', 'Нет выделения');
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