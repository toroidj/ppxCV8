//!*script パラメータで指定したコマンドの標準出力を返す
if ( PPx.Arguments.Count == 0 ) PPx.Quit(-1);
var Shell = PPx.CreateObject("WScript.Shell");
var exe = Shell.Exec(PPx.Arguments.Item(0));
if ( !exe.StdOut.AtEndOfStream ){
//	PPx.result = exe.StdOut.ReadLine(); // １行取得
	PPx.result = exe.StdOut.ReadAll(); // 全行取得
}
