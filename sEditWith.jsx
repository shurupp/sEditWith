#target InDesign;


var myDoc = app.activeDocument;

//GET SELECTION
var mySel = app.selection[0];
var myLink = '';
var myFileExt = '';
var phNameVer = '';
var illNameVer = '';
var phIDVer = 0;
var illIDVer = 0;

if (app.selection != 0) {

    //alert(app.selection[0].constructor.name, 'Constructor.name = ');

    /*
    LinkStatus.NORMAL
    LinkStatus.LINK_OUT_OF_DATE
    LinkStatus.LINK_MISSING
    LinkStatus.LINK_EMBEDDED
    LinkStatus.LINK_INACCESSIBLE
    */

    switch (app.selection[0].constructor.name) {
        case "Oval":
        case "Polygon":
        case "Rectangle":
            break;
        case "Image":
        case "PDF":
        case "EPS":
        case "AI":
            mySel.parent.select();
            mySel = app.selection[0];
            break;
        default:
            alert("Выбранный объект не содержит изображения");
            exit();
    }

    /*
    currentSel = mySel.allGraphics[0].Link;
    if (currentSel.isValid) {
        // do something here
        alert('Link valid');
    } else {
        alert('Нет связанного изображения', 'Error');
    }
*/

    try {
        myLink = mySel.allGraphics[0].itemLink;
    } catch (e) {
        alert('Нет связанного изображения', 'Error');
        exit();
    }

    //alert(myLink.status, 'myLink.status = ');
    //alert(myLink.linkType, 'myLink.linkType = ');

    if (myLink.status == LinkStatus.LINK_OUT_OF_DATE) {
        alert('Изображение не обновлено (LINK_OUT_OF_DATE)', 'Error');
        exit();
    }
    if (myLink.status == LinkStatus.LINK_EMBEDDED) {
        alert('Изображение внедрено (LINK_EMBEDDED)', 'Error');
        exit();
    }
    if (myLink.status == LinkStatus.LINK_MISSING) {
        alert('Изображение потеряно или недоступно (LINK_MISSING или LINK_INACCESSIBLE)', 'Error');
        exit();
    }
    if (myLink.status == LinkStatus.NORMAL) {
        myFilePath = myLink.filePath;
        myFileExt = myFilePath.split('.').reverse()[0];
    }

    myFileExt = myFileExt.toLowerCase();
    //myFileExt = mySel.allGraphics[0].itemLink.linkType
    //alert(myFileExt, 'Extension =');


    //var myListMenu = app.menus.item("$ID/Main").submenus.item("$ID/Edit").submenus.item("$ID/Edit With").submenus.everyItem().getElements();;

    var myListMenu = app.menuActions.everyItem().getElements();
    //var myListMenu = app.menus.everyItem().getElements();
    alert(myListMenu.length);
    for (var i = 0; i < myListMenu.length; i++) {
        if (myListMenu[i].name.search('Adobe Photoshop') != -1) {
            alert(myListMenu[i].area + ' : ' + myListMenu[i].name, '11111 ' + myListMenu[i].id);
            phNameVer = myListMenu[i].name;
            phIDVer = myListMenu[i].id;
            //alert(phNameVer, i + ' zzzzzzzz');
        }
        if (myListMenu[i].name.search('Adobe Illustrator') != -1) {
            alert(myListMenu[i].area + ' : ' + myListMenu[i].name, '11111 ' + myListMenu[i].id);
            illNameVer = myListMenu[i].name;
            illIDVer = myListMenu[i].id;
            //alert(IllNameVer, i + ' yyyyyyyyy');
        }
    }

    if (illNameVer == '') {
        alert('Adobe Illustrator не найден в контекстном меню', 'Error');
        //                exit();
    }

    if (phNameVer == '') {
        alert('Adobe Photoshop не найден в контекстном меню', 'Error');
        //              exit();
    }

    switch (myFileExt) {
        case "eps":
        case "ai":
        case "pdf":
        case "svg":
            //app.menuActions.itemByName(illNameVer).invoke();

            //alert(app.menuActions.itemByName(illNameVer), 'menu');
            //alert(app.menuActions.itemByName(illNameVer).name, 'name');

            //alert(app.menuActions.itemByID(illIDVer).id, 'id');
            //app.menuActions.itemByName(illNameVer).invoke();
            app.menuActions.itemByID(illIDVer).invoke();
            break;
        case "jpg":
        case "jpeg":
        case "jpe":
        case "bmp":
        case "png":
        case "tif":
        case "psd":
        case "webp":
        case "heic":
            //alert(app.menuActions.itemByName(phNameVer), 'menu');

            //alert(app.menuActions.itemByName(phNameVer).name, 'name = ');
            //alert (app.menuActions.itemByID(phIDVer).id, 'id = ');
            //app.menuActions.itemByName(phNameVer).invoke();
            app.menuActions.itemByID(phIDVer).invoke();
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

/* скрипт вывода всех menuaction
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
