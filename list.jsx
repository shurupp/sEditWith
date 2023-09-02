var arr = [];

for (var i = 0; i < app.menuActions.length; i++) {
    arr.push((i + 1) + " - " + app.menuActions[i].name);
}

var str = arr.join("\r");
WriteToFile(str);

function WriteToFile(text) {
    file = new File("~/Desktop/Menu actions.txt");
    file.encoding = "UTF-8";
    file.open("w");
    file.writeln(text); 
    file.close();
}