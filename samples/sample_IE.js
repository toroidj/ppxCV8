//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.report(
	"\r\n● " + PPx.ScriptName + "\r\n"+
	"Script Engine: " + PPx.ScriptEngineName + "/" + PPx.ScriptEngineVersion + "\r\n");

var IE = PPx.CreateObject("InternetExplorer.Application", "ie_");
try {
	IE._("trace"); // QuickJS のみ。イベントが発生したら全て PPx.report で通知
//	IE._("info"); // QuickJS のみ。メソッド・プロパティの一覧を表示。
//	IE._("event"); // QuickJS のみ。イベントの一覧を表示。
} catch(e) {};
IE.Visible = true;
IE.Navigate("http://toro.d.dooo.jp/");

	PPx.DisconnectObject(IE); // イベント発生を止める

function ie_OnVisible(visible){
	PPx.report("<<" + visible + (visible ? ":show>>" : ":hide>>"));
}
function ie_PropertyChange(prop){
	PPx.report("<<Change:" + prop + ">>");
}
function ie_DownloadComplete(){
	PPx.report("<<load Complete>>");
}
function ie_OnQuit(){
	PPx.report("<<close>>");
}
