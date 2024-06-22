//!*script
/* JScript / JScript(Chakra) / V8 / QuickJS 兼用

実行 [1回目-常駐開始][2回目-常駐解除][3回目-常駐開始][4回目-fire()]
(PPx以外から実行)  ppbw /c *script sample_resume.js,arg1,arg2 %: *script sample_resume.js,arg1,arg2 %: *script sample_resume.js,arg1,arg2%: *script ":fire",sample_resume.js,arg1,arg2
(PPxから実行)      ppbw /c %(*script sample_resume.js,arg1,arg2 %: *script sample_resume.js,arg1,arg2 %: *script sample_resume.js,arg1,arg2%: *script ":fire",sample_resume.js,arg1,arg2%)

複数インスタンスの例 [10001 1回目][10002 1回目][10003 1回目][10002 2回目-fire()][10002 3回目-fire()][10002 4回目-ppx_resume]
(PPx以外から実行)  ppbw /c *script :10001,sample_resume.js,arg1,arg2 %: *script :10002,sample_resume.js,arg1,arg2 %: *script :10003,sample_resume.js,arg1,arg2 %: *linemessage %*script(":10002,fire",sample_resume.js,arg1,arg2,arg3) %: *js ":10002","var exvalue = '**value**'; excall()" %: *script ":10002","",arg1,arg2,arg3
(PPxから実行)      ppbw /c %(*script :10001,sample_resume.js %: *script :10002,sample_resume.js %: *script :10003,sample_resume.js%: *linemessage %*script(":10002,fire",sample_resume.js,arg1,arg2,arg3) %: *script ":10002",sample_resume.js,arg1,arg2,arg3%)
*/

// グローバルスコープは、初めての時だけ実行される。

PPx.report("\r\n* sample_resume.js ("+ PPx.ScriptEngineName + " " + PPx.ScriptEngineVersion + ")\r\n");

PPx.report("start global\r\n");
var Count = 0;
var exvalue = ''; // 外部から変更される
ppx_resume(PPx.Argument(0), PPx.Argument(1));

// 常駐操作。コマンドラインで指定済みならその値を使用する
if ( PPx.StayMode < 2 ) PPx.StayMode = 2;

// ppx_resume() は、２回目以降の実行時に直接実行される。
// グローバルスコープの実行(start main の表示等)は行われない
function ppx_resume(a, b){
	PPx.report("function ppx_resume(" + a + ", " + b + "): StayMode = " + PPx.StayMode + "   Count = " + Count + "  exvalue = " + exvalue + "\r\n");
	Count++;
	if ( Count == 2 ) PPx.StayMode = 0; // ２回目実行の時は常駐解除
}

// ppx_finally() は、PPx.StayMode = 2 のとき、該当窓の終了時に呼び出される
// 親PPx の機能は全て使用不可能。メッセージを使用したいときは
// PPx.Echo のみ使用可能。
// ファイルを閉じる処理等の後始末を想定。
// 終了時に PPx を操作したい場合は、CLOSEEVENT を使用すること。
function ppx_finally(){
	PPx.Echo("function ppx_finally() instance = "+ PPx.StayMode + "  exvalue = " + exvalue);
}

// fire() は、スクリプトから呼ばれない。*scirpt ":fire" で実行される。
function fire(a, b, c, d, e){
	PPx.report("function fire(" + a + ", " + b + ", " + c + ", " + d + ", " + e + ") instance = " + PPx.StayMode + "  exvalue = " + exvalue + "\r\n");
	Count++;
	return "** fire() = " + Count + " **\r\n";
}

// excall() は、スクリプトから呼ばれない。*js ":n","exvalue = %C; excall();" で実行される。
function excall(){
	PPx.report("function excall() instance = " + PPx.StayMode + "  exvalue = " + exvalue + "\r\n");
}

// ここもグローバルスコープ
PPx.report("end global instance = "+ PPx.StayMode + "\r\n\r\n"); // 初めての時だけ実行
