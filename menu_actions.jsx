// Show menu actions and their ids
// Peter Kahrel
// See also http://kasyan.ho.com.ua/open_menu_item.html

// Example: app.menuActions.item('$ID/Find/Change...').invoke();


#targetengine session;

(function () {

	var hideJunk = true;
	var actions = [];
	var areas = [];
	var stripCode = / \([a-z]+?:\)/;  // Strip the occasional (xyz:) code
	var document = /\.(?:indd|jsx?(?:bin)?)$/i;
	// The following regex is English-specific,
	// can be changed to another language
	var junkArea = /^(?:Text Selection|Menu:Insert)/i;


	function thousand_sep (n) {
		return String (n).replace (/(\d)(?=(\d\d\d)+$)/g, '$1,')
	}


	function ignore (action) {
		if (!hideJunk) {
			return false;
		}
		return (action.id >= 57603 && action.id <= 61066)  // fonts and style names
					|| document.test (action.name)
					|| junkArea.test (action.area);
		}


	function sort_multi_column_list (key) {
		if (key === 'id') {
			actions.sort (function (a,b) {return a.id - b.id});
		} else if  (key === 'area') {
			actions.sort (function (a,b) {return a.area > b.area}); // Want to sort on name as well, but won't work
		} else {
			actions.sort (function (a,b) {return a.name > b.name});
		}

		list.removeAll();
		for (var i = 0; i < actions.length; i++) {
			with (list.add ('item', actions[i].name)) {
				subItems[0].text = actions[i].area;
				subItems[1].text = actions[i].id;
			}
		}
	}

	//-------------------------------------------------------------------------------------
	// We create two arrays here: the list and the Area dropdown.

	function create_lists () {
		var known = {};
		var a = app.menuActions.everyItem().getElements();
			
		for (var i = 0; i < a.length; i++) {
			//$.bp(actions[i].name.indexOf('(o:)')>=0);
			if (!ignore (a[i])) {
				actions.push ({
					uniqueName: a[i].name+'%%'+a[i].area,
					name: a[i].name.replace (stripCode,''),
					area: a[i].area,
					id: a[i].id,
				});
				if (!known[a[i].area]) {
					areas.push (a[i].area);
					known[a[i].area] = true;
				}
			}
		}
		// Initially sort by area
		actions.sort (function (x,y) {return x.area > y.area});
		areas.sort();
		areas.unshift('[All]'); // add [All] at the beginning of the array
	}
 
	//---------------------------------------------------------------------------------------------------------------
	// The interface

	create_lists();

	var column_widths = [200, 100, 100];
	var w = new Window ("palette {orientation: 'row', alignChildren: 'top'}");
	var total = 0;
	var dummy = w.add ('group');
	var list = dummy.add ('listbox', undefined, '', {multiselect: true, 
																	numberOfColumns: 3, 
																	showHeaders: true, 
																	columnTitles: ['Name', 'Area', 'ID'], 
																	//columnWidths: column_widths
																	});
	
	list.maximumSize.height = w.maximumSize.height-100;
	//	list.preferredSize.width = 800
		//list.preferredSize.height = w.maximumSize.height-100;
		
		
		var sidebar = w.add ('group {orientation: "column"}');
		
		var filter = sidebar.add ('group {orientation: "column", alignChildren: "right"}');
			var name_group = filter.add ('group');
				name_group.add ('statictext {text: "Name:"}');
				var search_name = name_group.add ('edittext {active: true}');
				
			//var keystring_group = filter.add ('group');
			//	keystring_group.add ('statictext {text: "Keystring:"}');
			//	var keystring_name = keystring_group.add ('edittext');

			var area_group = filter.add ('group');
				area_group.add ('statictext {text: "Area:"}');
				var area_dropdown = area_group.add ('dropdownlist', undefined, areas);
					area_dropdown.selection = 0;

			var id_group = filter.add ('group');
				id_group.add ('statictext {text: "ID:"}');
				var search_id = id_group.add ('edittext');
			
			var sort_group = filter.add ('group');
				sort_group.add ('statictext {text: "Sort:"}');
				var sort_dropdown = sort_group.add ('dropdownlist', undefined, ['Name', 'Area', 'ID']);
					sort_dropdown.selection = 1;
			
		var filter2 = sidebar.add ('group {orientation: "column", alignChildren: "left"}');
			var keyString = filter2.add ('button {text: "Show keystring"}');

		for (var i = filter.children.length-1; i >= 0; i--) {
			filter.children[i].children[1].preferredSize.width = 150;
		}
		// Populate the list
		
		for (var i = 0; i < actions.length; i++) {
			with (list.add ('item', actions[i].name)) {
				subItems[0].text = actions[i].area;
				subItems[1].text = actions[i].id;
			}
			total++
		}
		w.text += 'Menu actions (' + thousand_sep (total) + ' items)';


		search_name.onChange = function() {
			total = 0;
			var newlist = dummy.add ('listbox', list.bounds, '', {multiselect: true, 
																		numberOfColumns: 3, 
																		showHeaders: true, 
																		columnTitles: ['Name', 'Area', 'ID'], 
																	//	columnWidths: column_widths
																	});

			for (var i = 0; i < actions.length; i++) {
				if (actions[i].name.search(search_name.text) > -1) {
					with (newlist.add ('item', actions[i].name)) {
						subItems[0].text = actions[i].area;
						subItems[1].text = actions[i].id;
					}
					total++;
				}
			}
			dummy.remove (list);
			list = newlist;
			w.text = 'Menu actions (' + thousand_sep (total) + ' items)';
		}


		area_dropdown.onChange = function () {
			total = 0;
			var newlist = dummy.add ('listbox', list.bounds, '', {multiselect: true, 
																		numberOfColumns: 3, 
																		showHeaders: true, 
																		columnTitles: ['Name', 'Area', 'ID'], 
																	//	columnWidths: column_widths
																	});

			for (var i = 0; i < actions.length; i++) {
				if (actions[i].area.indexOf (area_dropdown.selection.text) == 0 || area_dropdown.selection.text == '[All]') {
					with (newlist.add ('item', actions[i].name)) {
						subItems[0].text = actions[i].area;
						subItems[1].text = actions[i].id;
						}
						total++;
					}
				}
			dummy.remove (list);
			list = newlist;
			w.text = 'Menu actions  (' + thousand_sep (total) + ' items)';
			}
		
		
		search_id.onChange = function () {
			list.selection = null;
			var L = list.items.length;
			for (var i = 0; i < L; i++) {
				if (list.items[i].subItems[1].text === search_id.text) {
					list.selection = i;
					break;
					}
				}
			}
		
		// From InDesign CC onDoubleClick doesn't work any longer
//~ 		list.onDoubleClick = function () {
//~ 			search_name.text = list.selection[0].text + '|' + action_object.list[list.selection[0].text].id;
//~ 			keystring_name.text = app.findKeyStrings(list.selection[0].text).join (' >> ');
//~ 		}

		keyString.onClick = function () {
			var s = list.selection[0].text;
			if (s !== '') {
				alert (app.findKeyStrings(s).join('\r'), 'Keystring: '+s, false);
			}
		}

		sort_dropdown.onChange = function () {
			var new_list = sort_multi_column_list (sort_dropdown.selection.text.toLowerCase());
		}
		
	w.show ();

}());