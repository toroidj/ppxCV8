//!*script
// QuickJS / V8 用
// QuickJS の場合、*scriptm のみ静的import対応。また、エラー・例外が返されず、
// 組み込みエラー通知が効かない為、try - cache で自前のエラー処理が必要。
import * as exmodule from "./sample_export.js";

PPx.report("\r\n* sample_import.js ("+ PPx.ScriptEngineName + " " + PPx.ScriptEngineVersion + ")\r\n");

try {
	import('./sample_export.js').then(module => {
		PPx.report("dynamic 10 * 123 = " + module.test(10) + "\r\n");
	});

	PPx.report("static 10 * 123 = " + exmodule.test(10) + "\r\n");
}catch(e){
	PPx.Echo("-- PPxQjs.dll error --\n" + PPx.ScriptName + "\n" + e + "\n" + e.stack);
}
