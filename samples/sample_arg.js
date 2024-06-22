//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.report("\r\n* sample_arg.js ("+ PPx.ScriptEngineName + " " + PPx.ScriptEngineVersion + ")\r\n");

var arg = PPx.Arguments;
var len = arg.length; // PPx.Arguments.length も可
var list;
if ( len == 0 ){
	PPx.report("no parameter\r\n");
}else{
	// 列挙の例1
	PPx.report("Sample 1: " + PPx.Argument(-1) + "\r\n");

	// 列挙の例2 (WSH/V8/QuickJS兼用)PPx.Argument(i), (WSH限定)PPx.Arguments(i)
	list = "Sample 2: ";
	for ( var i = 0; i < len ; i++ ){
		list = list + PPx.Argument(i) + ", ";
	}
	PPx.report(list + "\r\n");

	// 列挙の例3 (WSH/V8/QuickJS兼用)v.value, (WSH限定)v.Item (V8限定)v.Item()
	list = "Sample 3: ";
	for (var v = PPx.Arguments; !v.atEnd(); v.moveNext() ){
		list = list + v.value + ", ";
	}
	PPx.report(list + "\r\n");

	// 列挙の例4 (WSH/V8/QuickJS兼用)
	list = "Sample 4: ";
	for (var v = PPx.Enumerator(PPx.Arguments); !v.atEnd(); v.moveNext() ){
		list = list + v.item() + ", ";
	}
	PPx.report(list + "\r\n");

	// 列挙の例5 (V8/QuickJS 限定, Chakraでは結果が空欄)
	try {
		PPx.report("Sample 5: " + Array.from(PPx.Arguments).join(', ') + "\r\n");
	}catch(e){
		PPx.report("Sample 5: (JScript/9.DLL は使用不可)\r\n");
	}
	// 列挙の例2 (V8/QuickJS 限定)
	list = "Sample 6: ";
/*@cc_on @*/
/*@if (@_jscript)
	list += "(WSH では使用不可)";
  @else */
	try {
		for ( var v of PPx.Arguments ){
			list += v + ", ";
		}
	}catch(e){
		list += "(Chakra.DLL は使用不可)";
	}
/*@end @*/
	PPx.report(list + "\r\n");
}
