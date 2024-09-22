//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.WindowIDName = "1"; // (WSH版)C_Z -> CZaa に変更。V8/QuickJS版は常にCZaa形式

var text = PPx.ScriptEngineName;

var pane = PPx.Pane;
var tab = pane.Tab;

text += "\r\n\r\nGetComboItemCount = " + PPx.GetComboItemCount() + "\r\n"
	+ "Current Pane #" + pane.index + "(length = " + pane.length + ")\r\n"
	+ "Current Tab #" + tab.index + "(length = " + tab.length  + ")\r\n"
	+ "  IDName = " + tab.IDName + ", Name = " + tab.Name + "\r\n"
	+ "    Type = " + tab.Type + ", Lock = " + tab.Lock + ", Color/BackColor = " + tab.Color + "/" + tab.BackColor + "\r\n"

// (兼用)PPx.Pane.Tab.Item(0).Name (WSH限定)PPx.Pane.Tab(0).Name
text += "Tab 0--> Pane:"+PPx.Pane.Tab.item(0).Pane + " / " + PPx.Pane.Tab.item(0).Name + " / " + PPx.Pane.Tab.item(0).IDName;
text += "\r\nTab 1--> Pane:"+PPx.Pane.Tab.Item(1).Pane + " / " + PPx.Pane.Tab.Item(1).Name + " / " + PPx.Pane.Tab.item(1).IDName;


var items = PPx.Pane;
for ( ; !items.atEnd() ; items.moveNext() ){
	text += "\r\nPane.index = " + items.index + ", IDName = " + items.Tab.IDName;
}

var items = PPx.Pane.Tab;
for ( ; !items.atEnd() ; items.moveNext() ){
	text += "\r\nTab.index = " + items.index + ", IDName = " + items.IDName;
}

PPx.report(text);

PPx.report("\r\nTab 2: lock\r\n");
PPx.Pane.Tab.item(2).Lock = -1;

PPx.report("Tab current: change tab color\r\n");
var a = PPx.Pane;
a.Tab.Color = 0x808000;
a.Tab.BackColor = 0xffffff;

var b = a.Tab;
b.IndexFrom = "CB";
PPx.report("Tab IDName = CB, type = " + b.Type + "\r\n");
