//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.report("\r\n* sample_process.js ("+ PPx.ScriptEngineName + " " + PPx.ScriptEngineVersion + ")\r\n");

var WS = PPx.CreateObject("WbemScripting.SWbemLocator");
var WSL = WS.ConnectServer();
var items = WSL.ExecQuery("Select * From Win32_Process")

/*@cc_on @*/
/*@if (!@_jscript) */
if ( PPx.ScriptEngineName == "QuickJS" ){ // QuickJS 版のみ使用可能
	PPx.report("** for ( of ) ** (QuickJS only)\r\n");
	for ( item of items ){
		props = item.Properties_;
		for ( prop of props ){
			if ( prop.Name == "Name" ){
				PPx.report(prop.Value+"\r\n");
			}
		}
	}
}else
/*@end @*/
{	// 全種で使用可能。但し、V8 版は非常に遅い
	PPx.report("** Enumerator **\r\n");
	for (var item = PPx.Enumerator(items); !item.atEnd(); item.moveNext() ){
		var props = item.item().Properties_;
		for (var prop = PPx.Enumerator(props); !prop.atEnd(); prop.moveNext() ){
			if ( prop.item().Name == "Name" ){
				PPx.report(prop.item().Value+"\r\n");
			}
		}
	}
}
