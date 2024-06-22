//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.report(
	"\r\n● sample.js\r\n"+
	"Script Engine: " + PPx.ScriptEngineName + "\r\n" +
	"Engine Version: " + PPx.ScriptEngineVersion + "\r\n" +
	"ScriptFullName: " + PPx.ScriptFullName + "\r\n" +
	"ScriptName: " + PPx.ScriptName + "\r\n" +
	"WindowIDName: " + PPx.WindowIDName + "\r\n" +
	"EntryIndex: " + PPx.EntryIndex + "\r\n" +
	"EntryDisplayTop: " + PPx.EntryDisplayTop + "\r\n" +
	"EntryDisplayX: " + PPx.EntryDisplayX + "\r\n" +
	"EntryDisplayY: " + PPx.EntryDisplayY + "\r\n" +
	"WindowDirection: " + PPx.WindowDirection + "\r\n" +
	"Pane/Tab index: Pane=" + PPx.Pane.index + "/Tab=" + PPx.Pane.Tab.index + "\r\n" +
	"TabLock: " + PPx.Pane.Tab.Lock + "\r\n\r\n"
);

PPx.report("・カーソル移動による列挙サンプル");
if ( PPx.EntryFirstMark == 0 ){
	PPx.report("\r\nマークはありません");
}else{
	do {
		PPx.report("\r\n\r\nEntryIndex: " + PPx.EntryIndex + "\r\n" +
			"EntryName: " + PPx.EntryName + "\r\n" +
			"EntryComment: " + PPx.EntryComment + "\r\n" +
			"EntryAttributes: " + PPx.EntryAttributes + "\r\n" +
			"EntrySize: " + PPx.EntrySize + "\r\n" +
			"EntryMark: " + PPx.EntryMark
		);
	}while( PPx.EntryNextMark != 0 );
}

PPx.report("\r\n\r\n・PPx.Entry による、カーソル操作無しの列挙サンプル\r\n");
// マークがないときはカーソル上のエントリを対象とする
// マークの有無を判定するときは FirstMark が 0 であるかで判定する。
// (WSH)items.DateXXX (V8)items.DateXXX.ToString()
var items = PPx.Entry;
if ( items.FirstMark == 0 ){ PPx.report("マークはありません\r\n"); }
for ( ;; ){
	PPx.report("Index: " + items.index + "\r\n" +
		"Name: " + items.Name + "\r\n" +
		"ShortName: " + items.ShortName + "\r\n" +
		"Attributes: " + items.Attributes + "\r\n" +
		"Size: " + items.Size + "\r\n" +
		"Created: " + new Date(items.DateCreated) + "\r\n" +
		"LastModified: " + new Date(items.DateLastModified) + "\r\n" +
		"LastAccessed: " + new Date(items.DateLastAccessed) + "\r\n" +
		"State :" + items.State + "\r\n" +
		"Color :" + items.ExtColor + "\r\n" +
		"Comment: " + items.Comment + "\r\n" +
		"Comment 1: " + items.GetComment(1) + "\r\n" +
		"Comment 2: " + items.GetComment(2) + "\r\n" +
		"Comment 3: " + items.GetComment(3) + "\r\n\r\n"
	);
	if ( items.NextMark == 0 ) break;
}

// エントリ列挙の簡略記載例

// マークが無いときはカーソル位置を含める例1
PPx.report("・PPx.Entry 内蔵 Enumerator 使用例\r\n");
for (var a = PPx.Entry; !a.atEnd(); a.moveNext() ) PPx.report(a.Name + ", ");

// マークが無いときはカーソル位置を含める例2(V8版/QuickJS版で使用可能,WSH不可)
PPx.report("\r\n\r\n・PPx.Entry for ( of ) 使用例\r\n");
/*@cc_on @*/
/*@if (@_jscript)
	PPx.report("(WSH では使用不可)");
  @else */
	try {
		for ( a of PPx.Entry ) PPx.report(a.Name + ", ");
	}catch(e){
		PPx.report("(Chakra.DLL は使用不可。PPc 以外では PPx.Entry がない)");
	}
/*@end @*/

PPx.report("\r\n\r\n・PPx.Entry マークが無いときは何もしない例\r\n");
for (var a = PPx.Entry.AllMark; !a.atEnd(); a.moveNext() ) PPx.report(a.Name + ", ");

PPx.report("\r\n\r\n・PPx.Entry 表示順で全エントリ列挙の例\r\n");
// 表示順で全エントリ列挙の例
for (var a = PPx.Entry.AllEntry; !a.atEnd(); a.moveNext() ) PPx.report(a.Name + ", ");
PPx.report("\r\n");
