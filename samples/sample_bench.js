//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

var result;
if ( typeof ScriptEngine == 'function' ){
	result = ScriptEngine() + '/' + [ScriptEngineMajorVersion(), ScriptEngineMinorVersion(), ScriptEngineBuildVersion()].join('.');
}else if ( typeof PPx.ScriptEngineName == 'string' ){
	result = PPx.ScriptEngineName + '/' + PPx.ScriptEngineVersion;
}else{
	result = "unknown engine(Chakra ?)";
}

function fib(i) {
  if(i == 0 || i == 1) return i;
  return fib(i-1) + fib(i-2);
}

var startAt = new Date();
var varf = fib(30);
PPx.report("sample_bench.js: " + result + "  fib(30) = " + varf + ", Time: " + (new Date() - startAt) + " ms\r\n");
