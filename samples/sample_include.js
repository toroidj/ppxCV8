//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.report("\r\n* sample_include.js ("+ PPx.ScriptEngineName + " " + PPx.ScriptEngineVersion + ")\r\n");

PPx.Include('./sample_i_source.js');
PPx.report("\r\nstatic 10 * 123 = " + test(10) + "\r\n");
