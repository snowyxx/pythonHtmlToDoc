// $Id : $
var TREE_FORMAT = [
//0. left position
	-40,
//1. top position
	0,
//2. show +/- buttons
	true,
//3. couple of button images (collapsed/expanded/blank)
	['script/icon_plus.gif', 'script/icon_minus.gif','script/spacer.gif'],
//4. size of images (width, height,ident for nodes w/o children)
	[16,16,16],
//5. show folder image
	true,
//6. folder images (closed/opened/document)
	[],
//7. size of images (width, height)
	[],
//8. indentation for each level [0/*first level*/, 8/*second*/, 16/*third*/,...]
	[0,8,16,24],
//9. tree background color ("" - transparent)
	"",
//10. default style for all nodes
	"clsNode",
//11. styles for each level of menu (default style will be used for undefined levels)
	[],//["clsNodeL0","clsNodeL1","clsNodeL2","clsNodeL3","clsNodeL4"],
//12. true if only one branch can be opened at same time
	false,
//13. item pagging and spacing
	[0,2],
/************** PRO EXTENSIONS ********************/
//14. draw explorer like tree ( identation will be ignored )
	true,
//15. Set of explorer images (folder, openfolder, page, minus, minusbottom, plus, plusbottom, line, join, joinbottom)
	["script/icon_plus.gif", "script/icon_minus.gif", "script/icon_book_open.gif", "script/spacer.gif", "script/spacer.gif", "script/spacer.gif", "script/spacer.gif", "script/spacer.gif", "script/spacer.gif", "script/spacer.gif"],
//16. Explorer images width/height
	[16,16],
//17. if true state will be saved in cookies
	true,
//18. if true - relative position will be used. (tree will be opened in place where init() was called)
	true,
//19. width and height of initial rectangle for relative positioning
	[250,200],
//20. resize background //works only under IE4+, NS6+ for relative positioning
	false,
//21. support bgcolor changing for selected node
	false,
//22. background color for non-selected and selected node
	[]
];
