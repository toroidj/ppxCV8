//!*script
// JScript / JScript(Chakra) / V8 / QuickJS 兼用

PPx.report(
	"\r\n● " + PPx.ScriptName + "\r\n"+
	"Script Engine: " + PPx.ScriptEngineName + "/" + PPx.ScriptEngineVersion + "\r\n");

var errmsg = '';

// Win9x の場合、ここで応答無しになることがある。原因不明
var fso = PPx.CreateObject("scripting.filesystemobject");

var text =
  'BuildPath(\'c:\\aaa\\bbb\',\'ddd\') : ' + fso.BuildPath('c:\\aaa\\bbb', 'ddd')
+ '\r\nGetAbsolutePathName(\'temp.ext\') : ' + fso.GetAbsolutePathName('temp.ext')
+ '\r\nGetBaseName(\'temp.ext\') : ' + fso.GetBaseName('temp.ext')
+ '\r\nGetDriveName(\'temp.ext\') : ' + fso.GetDriveName('temp.ext')
+ '\r\nGetDriveName(\'x:\\abc\\temp.ext\') : ' + fso.GetDriveName('x:\\abc\\temp.ext')
+ '\r\nGetExtensionName(\'x:\\abc\\temp.ext\') : ' + fso.GetExtensionName('x:\\abc\\temp.ext')
+ '\r\nGetFileName(\'x:\\abc\\temp.ext\') : ' + fso.GetFileName('x:\\abc\\temp.ext')
+ '\r\nGetParentFolderName(\'x:\\abc\\def\\temp.ext\') : ' + fso.GetParentFolderName('x:\\abc\\def\\temp.ext');

try {
	text +=
	'\r\nGetSpecialFolder  0(Windows) : ' + fso.GetSpecialFolder(0)
	+ '\r\nGetSpecialFolder  1(System)  : ' + fso.GetSpecialFolder(1)
	+ '\r\nGetSpecialFolder  2(%temp%)  : ' + fso.GetSpecialFolder(2);
}catch(e){
	errmsg += '\r\n** fso.GetSpecialFolder 使用不可(V8版は即値指定不可)';
}

//PPx.report('\r\n\r\nCreateFolder' + fso.CreateFolder('CreateFolder')); // 複数階層の一括作成はできない

var folder = fso.GetFolder('.');

text +=
  '\r\n\r\nfso.GetFolder(\'.\') (' + fso.GetAbsolutePathName('.') + ')'
+ '\r\nName : ' + folder.Name
+ '\r\nParentFolder : ' + folder.ParentFolder // v8 で未実装
+ '\r\nPath : ' + folder.Path
+ '\r\nShortName : ' + folder.ShortName // QuickJS emuole版で未実装
+ '\r\nShortPath : ' + folder.ShortPath // QuickJS emuole版で未実装
+ '\r\nSize : ' + folder.Size // QuickJS emuole版で未実装
+ '\r\nType : ' + folder.Type;

var files = folder.Files;

	// V8 / QuickJS 版で使用可能。WSH 版は記載不可
	// ※CV8版は日時を取得できない
	// ※QuickJS 32bit版でnew Dateを使うと異常終了する
	var enumlist = '';
/*@cc_on @*/
/*@if (@_jscript)
	enumlist = "\r\n** WSH では使用不可";
  @else */
	for(var v of files){
		enumlist +=
		'\r\n-----\r\nFile Name: ' + v.Name
		+ '\r\nParentFolder : ' + v.ParentFolder
		+ '\r\nPath : ' + v.Path
		+ '\r\nShortName : ' + v.ShortName
		+ '\r\nShortPath : ' + v.ShortPath
		+ '\r\nSize : ' + v.Size
		+ '\r\nAttributes : ' + v.Attributes
		+ '\r\nnew Date(DateCreated) : ' + NewDate(v.DateCreated)
		+ '\r\nnew Date(DateLastAccessed) : ' + NewDate(v.DateLastAccessed)
		+ '\r\nnew Date(DateLastModified) : ' + NewDate(v.DateLastModified)
		+ '\r\nType: ' + v.Type;
	}
/*@end @*/
	text += '\r\n\r\nfor(var a of files)' + enumlist;
try {
	var enumlist = '';
	for (var v = new Enumerator(files); !v.atEnd(); v.moveNext() ){
		var item;
		item = v.item();
		enumlist += '\r\n Name : ' + item.Name;
	}
	text += '\r\n\r\nnew Enumerator(files)' + enumlist;
}catch(e){
	errmsg += '\r\n** new Enumerator 使用不可(Chakra.DLL, V8, QuickJS では PPx.Enumerator を代わりに使用する)';
}

try {
	// ※V8版は日時を取得できない
	// ※QuickJS 32bit版でnew Dateを使うと異常終了する
	var enumlist = '';
	for (var v = PPx.Enumerator(files); !v.atEnd(); v.moveNext() ){
		var item;
		item = v.item();
		enumlist +=
		'\r\n-----\r\nFile Name: ' + item.Name
		+ '\r\nParentFolder : ' + item.ParentFolder
		+ '\r\nPath : ' + item.Path
		+ '\r\nShortName : ' + item.ShortName
		+ '\r\nShortPath : ' + item.ShortPath
		+ '\r\nSize : ' + item.Size
		+ '\r\nAttributes : ' + item.Attributes
		+ '\r\nnew Date(DateCreated) : ' + NewDate(item.DateCreated)
		+ '\r\nnew Date(DateLastAccessed) : ' + NewDate(item.DateLastAccessed)
		+ '\r\nnew Date(DateLastModified) : ' + NewDate(item.DateLastModified)
		+ '\r\nType: ' + item.Type;
	}
	text += '\r\n\r\nPPx.Enumerator(files)' + enumlist;
}catch(e){
	errmsg += '\r\n** PPx.Enumerator 使用不可(R20以前)';
}

try {
	var enumlist = '';
	for (var v = PPx.Enumerator(folder.SubFolders); !v.atEnd(); v.moveNext() ){
		var item;
		item = v.item();
		enumlist +=
		'\r\n-----\r\nFile Name: ' + item.Name
		+ '\r\nParentFolder: ' + item.ParentFolder
		+ '\r\nPath: ' + item.Path
		+ '\r\nShortName: ' + item.ShortName
		+ '\r\nShortPath: ' + item.ShortPath
		+ '\r\nSize: ' + item.Size
		+ '\r\nType: ' + item.Type;
	}
	text += '\r\n\r\nPPx.Enumerator(folder.SubFolders)' + enumlist;
}catch(e){
}

var list = new Array(3);
function listitem(name, data) {
  this.name = name;
  this.data = data;
}

list[0] = new listitem('name1', 'data1');
list[1] = new listitem('name2', 'data2');
list[2] = new listitem('name3', 'data3');

try {
	var enumlist = '';
	for (var v = PPx.Enumerator(list); !v.atEnd(); v.moveNext() ){
		var item = v.item();
		enumlist +=
			'\r\n-----\r\n name : ' + item.name
			+ '\r\n data : ' + item.data;
	}
	text += '\r\n\r\nPPx.Enumerator(Array)' + enumlist;
}catch(e){
	errmsg += ('\r\n** PPx.Enumerator(Array) 使用不可(JScript9.DLL,Chakra.DLL)');
}

// Win9x の場合、ここで応答無しになることがある。原因不明
var stream = PPx.CreateObject('ADODB.stream');
var testname = "!! test !!.txt";

try {
	stream.Open();
	stream.Type = 2;
	stream.Charset = 'UTF-8';
//	stream.LineSeparator = 10;

	try {
		stream.WriteText("WriteText line 1 ", 1);
		stream.WriteText("WriteText line 2 ", 1);
		stream.WriteText("WriteText line 3 ", 1);
	}catch(e){
		stream.Write("Write line 1 ");
		stream.Write("Write line 2 ");
		stream.Write("Write line 3 ");
	}
	stream.SaveToFile(testname, 2);
	stream.Close();

	stream.Open();
	stream.Type = 2;
	stream.Charset = 'UTF-8';
	stream.LoadFromFile(testname);
	text += '\r\n\r\nADODB.stream test\r\n' + stream.ReadText(-1);
	stream.Close();
}catch(e){
	errmsg += ('\r\n** V8版は、各種定数を即値で指定できないため、ADODB 使用不可');
}

// fso.DeleteFile(testname);

PPx.report(text);
PPx.report('\r\n' + errmsg + '\r\n');

// 32bit QuickJS では Date class を使うと異常終了するので回避するための関数
function NewDate(dateSN)
{
	if ( PPx.option("Date") != false ) return new Date(dateSN);
	return "***";
}
