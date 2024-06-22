//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.linemessage(PPx.ScriptEngineName);
PPx.WindowIDName = "1"; // (WSH版)C_Z -> CZaa に変更。V8/QuickJS版は常にCZaa形式

PPx.Echo("Total Items:" + PPx.GetComboItemCount() );

// (兼用)PPx.Pane.Tab.Item(0).Name (WSH限定)PPx.Pane.Tab(0).Name
PPx.Echo("Tab 0--> Pane:"+PPx.Pane.Tab.item(0).Pane + " / " + PPx.Pane.Tab.item(0).Name + " / " + PPx.Pane.Tab.item(0).IDName);
PPx.Echo("Tab 1--> Pane:"+PPx.Pane.Tab.Item(1).Pane + " / " + PPx.Pane.Tab.Item(1).Name + " / " + PPx.Pane.Tab.item(1).IDName);

PPx.Echo("current Pane : " + PPx.Pane.index + " / " + PPx.Pane.Tab.IDName);

var items = PPx.Pane;
for ( ; !items.atEnd() ; items.moveNext() ){
	PPx.Echo ("Pane : " + items.index + " / " + items.Tab.IDName);
}

var items = PPx.Pane.Tab;
for ( ; !items.atEnd() ; items.moveNext() ){
	PPx.Echo ("Tab : " + items.index + " / " + items.IDName);
}

PPx.report("Tab 2: lock\r\n");
PPx.Pane.Tab.item(2).Lock = -1;

PPx.report("Tab current: change tab color\r\n");
var a = PPx.Pane;
a.Tab.Color = 0x808000;
a.Tab.BackColor = 0xffffff;

var b = a.Tab;
b.IndexFrom = "CB";
PPx.Echo("Tab CB type:" + b.Type);
