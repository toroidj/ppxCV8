/*-----------------------------------------------------------------------------
	Paper Plane xUI  Clear Script V8 Module
-----------------------------------------------------------------------------*/
#include "pch.h"
#include <dispex.h>
#include "PPxCV8.h"
#pragma comment( lib, "Ole32.Lib" )

using namespace System::IO;

namespace ComTypes = System::Runtime::InteropServices::ComTypes;
namespace InteropServices = System::Runtime::InteropServices;

ref class CPPxObjects;
ref class CInstance;
int RunScript(PPXAPPINFOW* ppxa, PPXMCOMMANDSTRUCT* pxc, int file);
String^ GetScriptText(String^ param, int file);
void FreeStayInstance(void);
void DropStayInstance(CInstance^ instance);

long PPxVersion = 0;

// PPx.setValue / PPx.getValue の保存先
gcroot<Collections::Concurrent::ConcurrentDictionary<String^, Object^>^> g_value = nullptr;
#define gc_value static_cast<Collections::Concurrent::ConcurrentDictionary<String^, Object^>^>(g_value)

gcroot<Collections::Generic::List<CInstance^>^> g_StayInstance = nullptr;
#define gc_StayInstance static_cast<Collections::Generic::List<CInstance^>^>(g_StayInstance)

// pplibがあるディレクトリ。インライン実行時の import カレントに使用。
gcroot <String^> g_PPxPath = nullptr;

// Instance value(CInstance が管理)
typedef struct {
	PPXAPPINFOW* ppxa;
	PPXMCOMMANDSTRUCT* pxc;
	int ModuleMode; // Javascript module として扱うかどうか

	struct {
		PPXAPPINFOW *ppxa; // 非アクティブ時に使用する ppxa 要 HeapFree
		HWND hWnd;
		#define ScriptStay_None		0
		#define ScriptStay_Cache	1
		#define ScriptStay_Stay		2
		// ID 0-9999 setting
		// ID 2000-999999999 user
		#define ScriptStay_FirstUserID 2000
		#define ScriptStay_FirstAutoID 0x40000000
		#define ScriptStay_MaxAutoID 0x7fffff00
		int mode;

		DWORD threadID;
		int entry; // 使用中カウンタ
	} stay;
} InstanceValueStruct;

typedef struct {
	PPXAPPINFOW* ppxa;
	PPXMCOMMANDSTRUCT* pxc;
} OldPPxInfoStruct;

DWORD StayInstanseIDserver = ScriptStay_FirstAutoID;


void PopupMessage(PPXAPPINFOW* ppxa, const WCHAR* message)
{
	if ( (ppxa == NULL) ||
		 (ppxa->Function == NULL) ||
		 (ppxa->Function(ppxa, PPXCMDID_MESSAGE, (void *)message) == 0) ){
		HWND hWnd = (ppxa == NULL) ? NULL : ppxa->hWnd;
		MessageBoxW(hWnd, message, L"PPxCV8.dll", MB_OK);
	}
}

void PopupMessage(PPXAPPINFOW* ppxa, String^ message)
{
	marshal_context ctx;
	::PopupMessage(ppxa, ctx.marshal_as<const wchar_t*>(message));
}

BOOL TryOpenClipboard(HWND hWnd)
{
	int trycount = 6;

	for (;;) {
		if (IsTrue(OpenClipboard(hWnd))) return TRUE;
		if (--trycount == 0) return FALSE;
		Sleep(20);
	}
}

#if 0
BOOL GetRegStringW(HKEY hKey, const WCHAR* path, const WCHAR* name, WCHAR* dest, DWORD size)
{
	HKEY HK;
	DWORD t, s;

	if (::RegOpenKeyExW(hKey, path, 0, KEY_READ, &HK) == ERROR_SUCCESS) {
		s = size;
		if (::RegQueryValueExW(HK, name, NULL, &t, (LPBYTE)dest, &s)
			== ERROR_SUCCESS) {
			::RegCloseKey(HK);
			return TRUE;
		}
		::RegCloseKey(HK);
	}
	return FALSE;
}

void Debug_DispIID(REFIID riid)
{
	WCHAR* iidstring, path[MAX_PATH], name[MAX_PATH];

	::StringFromIID(riid, &iidstring);
	wsprintfW(path, L"Interface\\%s", iidstring);
	if (GetRegStringW(HKEY_CLASSES_ROOT, path, L"", name, sizeof(name))) {
		MessageW(name);
	}
	else {
		MessageW(path);
	}
}
#endif

#define DISPID_SENDEVENT 0
[ComVisible(true)]
public ref class CEventProxy
{
public:
	CEventProxy(V8ScriptEngine^ engine) {
		m_engine = engine;
	}

	#define IVPARAM 8
	[DispId(DISPID_SENDEVENT)]
	void SendEvent(String^ name, int argc,
			Object^ param1, Object^ param2, Object^ param3, Object^ param4,
			Object^ param5, Object^ param6, Object^ param7, Object^ param8)
	{
		marshal_context ctx;

		array<Object^>^ args = gcnew array<Object^>(argc);
		if ( argc >= 1 ) args[0] = param1;
		if ( argc >= 2 ) args[1] = param2;
		if ( argc >= 3 ) args[2] = param3;
		if ( argc >= 4 ) args[3] = param4;
		if ( argc >= 5 ) args[4] = param5;
		if ( argc >= 6 ) args[5] = param6;
		if ( argc >= 7 ) args[6] = param7;
		if ( argc >= 8 ) args[7] = param8;
		m_engine->Invoke(name, args);
		delete args;
	}

private:
	V8ScriptEngine^ m_engine;
};

public class CEventSink : public IDispatch
{
private:
	int m_refCount;
	PPXAPPINFOW** m_ppxa;
	IDispatch* m_EventProxy; // 親が生存管理する
	ITypeInfo* m_EventType; // 自身で生存管理する
	WCHAR m_prefix[256];

public:
	CEventSink(IDispatch* EventProxy, ITypeInfo* pEventType, const wchar_t* prefix) {
		m_refCount = 1; // addref 1回分
		m_EventProxy = EventProxy;
		m_EventProxy->AddRef();
		m_EventType = pEventType;
		wcscpy(m_prefix, prefix);
	}
	virtual ~CEventSink() {};
	//----------------------------------------------------- IUnknown
	STDMETHODIMP_(ULONG) AddRef() {
		return ++m_refCount;
	}
	STDMETHODIMP_(ULONG) Release() {
		if (--m_refCount == 0) {
			m_EventType->Release();
			delete this;
			return 0;
		}
		return m_refCount;
	}

	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) {
		if ((riid == IID_IUnknown) || (riid == IID_IDispatch)) {
			*ppvObj = this;
			AddRef();
			return S_OK;
		}
		*ppvObj = NULL;
		return E_NOINTERFACE;
	}
	//-------------------------------------------------- IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT* iTInfo) {
		*iTInfo = 0;
		return S_OK;
	}
	STDMETHODIMP GetTypeInfo(UINT, LCID, ::ITypeInfo** ppTInfo) {
		*ppTInfo = NULL;
		return DISP_E_BADINDEX;
	}
	STDMETHODIMP GetIDsOfNames(REFIID riid, OLECHAR** rgszNames, UINT cNames, LCID, DISPID* rgDispId) {
		return E_NOTIMPL;
	}

	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lc, WORD wFlags, ::DISPPARAMS* pDispParams, VARIANT* pVarResult, ::EXCEPINFO* ei, UINT* puArgErr) {
		WCHAR dispname[300];

		if (pVarResult != nullptr) {
			pVarResult->vt = VT_NULL;
		}

		dispname[0] = '\0';
		if (m_EventType != nullptr) {
			UINT c;
			BSTR name;

			if (SUCCEEDED(m_EventType->GetNames(dispIdMember, &name, 1, &c))) {
				wsprintfW(dispname, L"%s%s", m_prefix, name);
			}
			::SysFreeString(name);
		}
		if (dispname[0] == '\0') {
			wsprintfW(dispname, L"%s%d", m_prefix, dispIdMember);
		}

		if (m_EventProxy != nullptr) {
			::DISPPARAMS dispparam;

			VARIANT *input = new VARIANT[2 + IVPARAM], nullvar;

			::VariantInit(&nullvar);
			input[IVPARAM + 1].vt = VT_BSTR;
			input[IVPARAM + 1].bstrVal = ::SysAllocString(dispname);
			input[IVPARAM].vt = VT_I4;
			input[IVPARAM].intVal = pDispParams->cArgs;
/*
if ( pDispParams->cArgs ) {
	WCHAR a[100];
	wsprintfW(a, L"%s %d %d",dispname,  pDispParams->rgvarg[0].vt, pDispParams->rgvarg[0].boolVal);
	MessageW(a);
}
*/
			DWORD cargs = pDispParams->cArgs;
			if ( cargs >= IVPARAM ) cargs = IVPARAM;

			for ( DWORD i = 0; i < IVPARAM; i++ ){
				input[IVPARAM - 1 - i] = ( i < cargs ) ?
					pDispParams->rgvarg[pDispParams->cArgs - 1 - i] : nullvar;
			}

			dispparam.rgvarg = input;
			dispparam.cArgs = 2 + IVPARAM;
			dispparam.cNamedArgs = 0;
			dispparam.rgdispidNamedArgs = NULL;

			HRESULT hr = m_EventProxy->Invoke(DISPID_SENDEVENT, IID_NULL, lc, DISPATCH_METHOD, &dispparam, NULL, ei, puArgErr);

			::VariantClear(&input[0]);
			delete[] input;
// Messagef("%x",hr);
			return hr;
		}
		else {
			return E_FAIL;
		}
	}
};

public ref class CEventInstance
{
private:
	CEventProxy^ gcEventProxy;
	IntPtr gcProxyDispatch;
	IDispatch* pProxyDispatch;
	::IConnectionPoint* pPoint;
	DWORD m_connectid;

public:
	Object^ SourceObject;

	CEventInstance(V8ScriptEngine^ engine) {
		gcEventProxy = gcnew CEventProxy(engine);
		gcProxyDispatch = Marshal::GetIDispatchForObject(gcEventProxy);
		pProxyDispatch = static_cast<IDispatch*>(gcProxyDispatch.ToPointer());
		pPoint = NULL;
	}

	~CEventInstance() {
		this->!CEventInstance();
	}
	!CEventInstance() {
		Disconnect();
		Marshal::Release(gcProxyDispatch);
		delete gcEventProxy;
	}

	HRESULT Connect(Object^ gcSourceObject, ::IConnectionPoint* Point, ::ITypeInfo* pEventType, String^ prefix) {
		marshal_context ctx;
		DWORD connectid;
		HRESULT result;
		CEventSink* EventSink;

		SourceObject = gcSourceObject;
		pPoint = Point;
		EventSink = new CEventSink(pProxyDispatch, pEventType, ctx.marshal_as<const wchar_t*>(prefix));
		result = pPoint->Advise(EventSink, &connectid);
		EventSink->Release();
		m_connectid = connectid;
		return result;
	}

	void Disconnect(void) {
		if (pPoint == NULL) return;
		pPoint->Unadvise(m_connectid);
		pPoint->Release();
		pPoint = NULL;
	}
};

String^ GetArgumentsItem(PPXAPPINFOW* ppxa, PPXMCOMMANDSTRUCT* pxc, int nIndex)
{
	const WCHAR* ptr;
	int i;

	if ((nIndex >= 0) && (nIndex < ((int)pxc->paramcount - 1))) {
		ptr = pxc->param;
		for (i = -1; i < nIndex; i++) {
			ptr += wcslen(ptr) + 1;
		}
		return marshal_as<String^>(ptr);
	}

	if (nIndex == -1) {
		BSTR param;
		String^ s_result;
		const WCHAR* ptr;

		param = (BSTR)ppxa->Function(ppxa, PPXCMDID_GETRAWPARAM, NULL);
		if (param == NULL) return "";
		// 1つ目(ファイル・スクリプト)をスキップ
		ptr = param;
		while ((*ptr == ' ') || (*ptr == '\t')) ptr++;
		if (*ptr == '\"') {
			ptr++;
			for (;;) {
				if (*ptr == '\0') break;
				if (*ptr++ == '\"') break;
			}
			while ((*ptr == ' ') || (*ptr == '\t')) ptr++;
		}
		else {
			for (;;) {
				if ((*ptr == '\0') || (*ptr == ',')) break;
				ptr++;
			}
		}
		if (*ptr == ',') ptr++;
		s_result = marshal_as<String^>(ptr);
		::SysFreeString(param);
		return s_result;
	}
	return "";
}

public ref class CArguments : public IEnumerable
{
private:
	int m_index;
	PPXAPPINFOW** m_ppxa;
	PPXMCOMMANDSTRUCT** m_pxc;

	ref class CArgumentsEnum : public IEnumerator
	{
	private:
		PPXAPPINFOW** m_ppxa;
		PPXMCOMMANDSTRUCT** m_pxc;
		int m_index;

	public:
		CArgumentsEnum(PPXAPPINFOW** ppxa, PPXMCOMMANDSTRUCT** pxc)
		{
			m_ppxa = ppxa;
			m_pxc = pxc;
			m_index = -1;
		}

		virtual property Object^ Current
		{
			Object^ get()
			{
				return GetArgumentsItem(*m_ppxa, *m_pxc, m_index);
			}
		}

		virtual void Reset()
		{
			m_index = -1;
		}

		virtual bool MoveNext()
		{
			m_index++;
			if (m_index < ((int)(*m_pxc)->paramcount - 1)) {
				return true;
			}
			else {
				return false;
			}
		}
	};

public:
	CArguments(PPXAPPINFOW** ppxa, PPXMCOMMANDSTRUCT** pxc) {
		m_ppxa = ppxa;
		m_pxc = pxc;
		m_index = 0;
	};

	virtual IEnumerator^ GetEnumerator()
	{
		return gcnew CArgumentsEnum(m_ppxa, m_pxc);
	}

	String^ Item()
	{
		return GetArgumentsItem(*m_ppxa, *m_pxc, m_index);
	}

	String^ Item(int nIndex)
	{
		return GetArgumentsItem(*m_ppxa, *m_pxc, nIndex);
	}

	String^ item(int nIndex)
	{
		return GetArgumentsItem(*m_ppxa, *m_pxc, nIndex);
	}

	property String^ value
	{
		String^ get() {
			return GetArgumentsItem(*m_ppxa, *m_pxc, m_index);
		}
	}

	property int length
	{
		int get() {
			return (*m_pxc)->paramcount - 1;
		}
	}
	property int Count { int get() { return length; } }

	bool atEnd()
	{
		if (m_index >= static_cast<int>((*m_pxc)->paramcount - 1)) {
			return 1;
		}
		else {
			return 0;
		}
	}

	void moveNext()
	{
		if (m_index < static_cast<int>((*m_pxc)->paramcount - 1)) {
			m_index++;
		}
	}
};

public ref class CTab : public IEnumerable
{
private:
	PPXAPPINFOW** m_ppxa;
	int m_PaneIndex;
	int m_TabIndex;

	ref class CTabEnum : public IEnumerator
	{
	private:
		PPXAPPINFOW** m_ppxa;
		int m_PaneIndex;
		int m_TabIndex;

	public:
		CTabEnum(PPXAPPINFOW** ppxa, int paneindex)
		{
			m_ppxa = ppxa;
			m_PaneIndex = paneindex;
			Reset();
		}

		virtual property Object^ Current
		{
			Object^ get()
			{
				if (m_TabIndex < 0) return nullptr;
				return gcnew CTab(m_ppxa, m_PaneIndex, m_TabIndex);
			}
		}

		virtual void Reset()
		{
			m_TabIndex = -1;
		}

		virtual bool MoveNext()
		{
			int maxvalue = m_PaneIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABCOUNT, &maxvalue);

			if ((m_TabIndex + 1) >= maxvalue) return false;
			m_TabIndex++;
			return true;
		}
	};

public:
	CTab(PPXAPPINFOW** ppxa, int paneindex)
	{
		m_ppxa = ppxa;
		m_PaneIndex = paneindex;
		m_TabIndex = -1;
	};

	CTab(PPXAPPINFOW** ppxa, int paneindex, int tabindex)
	{
		m_ppxa = ppxa;
		m_PaneIndex = paneindex;
		m_TabIndex = tabindex;
	};

	virtual IEnumerator^ GetEnumerator()
	{
		return gcnew CTabEnum(m_ppxa, m_PaneIndex);
	}

	property int length
	{
		int get() {
			int value = m_PaneIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABCOUNT, &value);
			return value;
		}
	}
	property int Count { int get() { return length; } }

	property String^ Name
	{
		String^ get() {
			DWORD buf[CMDLINESIZE / sizeof(DWORD)];

			buf[0] = m_PaneIndex;
			buf[1] = m_TabIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABNAME, &buf);

			return marshal_as<String^>((WCHAR*)&buf);
		}
		void set(String^ name) {
			marshal_context ctx;
			DWORD buf[4];

			buf[0] = m_PaneIndex;
			buf[1] = m_TabIndex;
			*(const wchar_t**)&buf[2] = ctx.marshal_as<const wchar_t*>(name);
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_SETTABNAME, &buf);
		}
	}

	property String^ IDName
	{
		String^ get() {
			DWORD buf[REGEXTIDSIZE / (sizeof(DWORD) / sizeof(WCHAR))];

			buf[0] = m_PaneIndex;
			buf[1] = m_TabIndex;
			if ((*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABLONGID, &buf) == 0) {
				buf[0] = m_PaneIndex;
				buf[1] = m_TabIndex;
				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABIDNAME, &buf);
				if (((WCHAR*)&buf)[1] == '_') {
					memmove(((WCHAR*)&buf) + 1, ((WCHAR*)&buf) + 2, (strlenW(((WCHAR*)&buf) + 2) + 1) * 2);
				}
			}
			return marshal_as<String^>((WCHAR*)&buf);
		}
	}

	property int Lock
	{
		int get()
		{
			int tmp[2];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_DIRLOCK, &tmp);
			return tmp[0];
		}

		void set(int value)
		{
			int tmp[4];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			tmp[2] = value;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_SETDIRLOCK, &tmp);
		}
	}

	property int index
	{
		int get()
		{
			int tmp[2];

			if (m_TabIndex >= 0) {
				return m_TabIndex;
			}
			else {
				tmp[0] = m_PaneIndex;
				tmp[1] = m_TabIndex;
				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABINDEX, &tmp);
				return tmp[0];
			}
		}

		void set(int value)
		{
			m_TabIndex = value;
		}
	}

	property int Index {
		int get() { return index; }
		void set(int num) { index = num; }
	}

	CTab^ Item()
	{
		CTab^ tab = gcnew CTab(m_ppxa, m_PaneIndex);
		tab->index = m_TabIndex;
		return tab;
	}

	CTab^ Item(int index)
	{
		CTab^ tab = gcnew CTab(m_ppxa, m_PaneIndex);
		tab->index = index;
		return tab;
	}

	CTab^ item(int index)
	{
		CTab^ tab = gcnew CTab(m_ppxa, m_PaneIndex);
		tab->index = index;
		return tab;
	}

	property int Color
	{
		int get()
		{
			int tmp[2];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_TABTEXTCOLOR, &tmp);
			return tmp[0];
		}

		void set(int value)
		{
			int tmp[4];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			tmp[2] = value;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_SETTABTEXTCOLOR, &tmp);
		}
	}

	property int BackColor
	{
		int get()
		{
			int tmp[2];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_TABBACKCOLOR, &tmp);
			return tmp[0];
		}

		void set(int value)
		{
			int tmp[4];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			tmp[2] = value;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_SETTABBACKCOLOR, &tmp);
		}
	}

	property int Type
	{
		int get()
		{
			int tmp[2];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOWNDTYPE, &tmp);
			return tmp[0];
		}

		void set(int value)
		{
			int tmp[4];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			tmp[2] = value;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_SETCOMBOWNDTYPE, &tmp);
		}
	}

	property int Pane
	{
		int get()
		{
			int tmp[2];

			if (m_PaneIndex >= 0) {
				return m_PaneIndex;
			}
			else {
				tmp[0] = m_PaneIndex;
				tmp[1] = m_TabIndex;
				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABPANE, &tmp);
				return tmp[0];
			}
		}

		void set(int value)
		{
			int tmp[4];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			tmp[2] = value;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_SETCOMBOTABPANE, &tmp);
			m_PaneIndex = value;
		}
	}

	property int TotalCount
	{
		int get()
		{
			int tmp[2];

			tmp[0] = m_PaneIndex;
			tmp[1] = m_TabIndex;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABCOUNT, &tmp);
			return tmp[0];
		}
	}

	property int TotalPPcCount
	{
		int get()
		{
			return 0;
		}
	}

	int IndexFrom(String^ str)
	{
		marshal_context ctx;
		PPXUPTR_TABINDEXSTRW tmp;
		DWORD_PTR result;

		tmp.pane = m_PaneIndex;
		tmp.tab = m_TabIndex;
		tmp.str = (wchar_t*)ctx.marshal_as<const wchar_t*>(str);
		result = (*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOGETTAB, &tmp);
		m_PaneIndex = tmp.pane;
		m_TabIndex = tmp.tab;
		return (result == PPXA_NO_ERROR);
	}

	int Execute(String^ str)
	{
		marshal_context ctx;
		int resultcode;
		PPXUPTR_TABINDEXSTRW tmp;

		tmp.pane = m_PaneIndex;
		tmp.tab = m_TabIndex;
		tmp.str = (wchar_t*)ctx.marshal_as<const wchar_t*>(str);
		resultcode = static_cast<int>((*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABEXECUTE, &tmp));
		if (resultcode == 1) resultcode = NO_ERROR;
		return resultcode;
	}

	String^ Extract(String^ str)
	{
		WCHAR* bufw;
		marshal_context ctx;
		const WCHAR* param;
		PPXUPTR_TABINDEXSTRW tmp;
		LONG_PTR long_result;
		size_t paramlen;
		String^ s_result;

		// ※ param = marshal_as<std::wstring>(str).c_str(); だと失敗するので分割 2022-07
		param = ctx.marshal_as<const wchar_t*>(str);
		paramlen = wcslen(param) + 1;
		if (paramlen < CMDLINESIZE) paramlen = CMDLINESIZE;
		bufw = new WCHAR[paramlen];

		tmp.pane = m_PaneIndex;
		tmp.tab = m_TabIndex;
		wcscpy(bufw, param);
		tmp.str = bufw;
		long_result = (*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABEXTRACTLONG, &tmp);
		if (long_result >= 0x10000) {
			BSTR b_result;

			b_result = (BSTR)long_result;
			if (b_result == NULL) {
				s_result = "";
			}
			else {
				s_result = marshal_as<String^>(b_result);
				::SysFreeString(b_result);
			}
			delete[] bufw;
			return s_result;
		}

		tmp.pane = m_PaneIndex;
		tmp.tab = m_TabIndex;
		tmp.str = bufw;
		wcscpy(bufw, param);
		(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABEXTRACT, &tmp);
		s_result = marshal_as<String^>(bufw);
		delete[] bufw;
		return s_result;
	}

	bool atEnd()
	{
		int maxtabs = 0;

		if (m_TabIndex < 0) m_TabIndex = 0;
		maxtabs = m_PaneIndex;
		(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOTABCOUNT, &maxtabs);
		if (m_TabIndex >= maxtabs) {
			return 1;
		}
		else {
			return 0;
		}
	}

	void moveNext()
	{
		m_TabIndex++;
	}
};

public ref class CPane : public IEnumerable
{
private:
	PPXAPPINFOW** m_ppxa;
	int m_PaneIndex; // -1:実行元の窓

	ref class CPaneEnum : public IEnumerator
	{
	private:
		PPXAPPINFOW** m_ppxa;
		int m_PaneIndex;

	public:
		CPaneEnum(PPXAPPINFOW** ppxa)
		{
			m_ppxa = ppxa;
			Reset();
		}

		virtual property Object^ Current
		{
			Object^ get()
			{
				if (m_PaneIndex < 0) return nullptr;
				return gcnew CPane(m_ppxa, m_PaneIndex);
			}
		}

		virtual void Reset()
		{
			m_PaneIndex = -1;
		}

		virtual bool MoveNext()
		{
			int maxvalue;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOSHOWPANES, &maxvalue);

			if ((m_PaneIndex + 1) >= maxvalue) return false;
			m_PaneIndex++;
			return true;
		}
	};

public:
	CPane(PPXAPPINFOW** ppxa, int paneindex)
	{
		m_ppxa = ppxa;
		m_PaneIndex = paneindex;
	};

	virtual IEnumerator^ GetEnumerator()
	{
		return gcnew CPaneEnum(m_ppxa);
	}

	property int length
	{
		int get() {
			int value;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOSHOWPANES, &value);
			return value;
		}
	}
	property int Count { int get() { return length; } }

	property int index
	{
		int get() {
			if (m_PaneIndex >= 0) {
				return m_PaneIndex;
			}
			else {
				int tmp[2];

				tmp[0] = m_PaneIndex;
				tmp[1] = 0; // 未使用であるが、念のため
				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOSHOWINDEX, &tmp);
				return tmp[0];
			}
		}
		void set(int index) {
			m_PaneIndex = index;
		}
	}
	property int Index {
		int get() { return index; }
		void set(int num) { index = num; }
	}

	int IndexFrom(String^ str)
	{
		marshal_context ctx;

		m_PaneIndex = static_cast<int>((*m_ppxa)->Function(*m_ppxa,
			PPXCMDID_COMBOGETPANE,
			(void*)ctx.marshal_as<const wchar_t*>(str)));
		return (m_PaneIndex >= 0);
	}

	CPane^ Item()
	{
		CPane^ pane = gcnew CPane(m_ppxa, m_PaneIndex);
		return pane;
	}

	CPane^ Item(int index)
	{
		CPane^ pane = gcnew CPane(m_ppxa, index);
		return pane;
	}

	CPane^ item(int index)
	{
		CPane^ pane = gcnew CPane(m_ppxa, index);
		return pane;
	}

	property CTab^ Tab
	{
		CTab^ get() {
			CTab^ tab = gcnew CTab(m_ppxa, m_PaneIndex);
			tab->index = index;
			return tab;
		}
	}

	property CTab^ tab
	{
		CTab^ get() {
			CTab^ tab = gcnew CTab(m_ppxa, m_PaneIndex);
			tab->index = index;
			return tab;
		}
	}

	bool atEnd()
	{
		int maxpanes = 0;

		if (m_PaneIndex < 0) m_PaneIndex = 0;
		(*m_ppxa)->Function(*m_ppxa, PPXCMDID_COMBOSHOWPANES, &maxpanes);
		if (m_PaneIndex >= maxpanes) {
			return 1;
		}
		else {
			return 0;
		}
	}

	void moveNext()
	{
		m_PaneIndex++;
	}
};

bool IsEntryVailed(PPXAPPINFOW* ppxa, int newindex)
{
	int state;
	PPXUPTR_INDEX_UPATHW tmp;

	state = newindex;
	ppxa->Function(ppxa, PPXCMDID_ENTRYSTATE, &state);
	if ((state & 7) < 2) return false; // < ECS_NORMAL

	tmp.index = newindex;
	ppxa->Function(ppxa, PPXCMDID_ENTRYNAME, &tmp);
	if ((tmp.path[0] == '.') &&
		((tmp.path[1] == '\0') ||
			((tmp.path[1] == '.') && (tmp.path[2] == '\0')))) { // . or ..
		return false;
	}
	return true;
}

public ref class CEntry : public IEnumerable
{
private:
	PPXAPPINFOW** m_ppxa;
	int m_index, m_enum;

	ref class CEntryEnum : public IEnumerator
	{
	private:
		PPXAPPINFOW** m_ppxa;
		int m_index;
		int m_mode;

		bool CheckNextEntry(int newindex)
		{
			int maxindex;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_DIRTTOTAL, &maxindex);
			for (;;) {
				if (newindex >= maxindex) return false;
				if (IsEntryVailed(*m_ppxa, newindex) == true) {
					m_index = newindex;
					return true;
				}
				newindex++;
			}
		}

	public:
		CEntryEnum(PPXAPPINFOW** ppxa, int mode)
		{
			m_ppxa = ppxa;
			m_mode = mode;
			Reset();
		}

		virtual property Object^ Current
		{
			Object^ get()
			{
				if (m_index == -1) return nullptr;
				return gcnew CEntry(m_ppxa, m_index);
			}
		}

		virtual void Reset()
		{
			m_index = -1;
		}

		virtual bool MoveNext()
		{
			int result;

			if (m_index == -1) { // 初めて
				if (m_mode == 2) return CheckNextEntry(0);

				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKFIRST_HS, &result);
				if (result == -1) {
					if (m_mode == 0) {
						(*m_ppxa)->Function(*m_ppxa, PPXCMDID_CSRINDEX, &result);
						m_index = result;
					}
					else {
						return false;
					}
				}
				else {
					m_index = result;
				}
				if (IsEntryVailed(*m_ppxa, m_index) == true) return true;
			}

			// ２回目以降
			if (m_mode == 2) return CheckNextEntry(m_index + 1);
			for (;;) {
				result = m_index;
				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKNEXT_HS, &result);
				if ((result == -1) || (result == m_index)) return false;
				m_index = result;
				if (IsEntryVailed(*m_ppxa, result) == true) return true;
			}
		}
	};

	bool CheckNextEntryAt(int newindex)
	{
		int maxindex;

		(*m_ppxa)->Function(*m_ppxa, PPXCMDID_DIRTTOTAL, &maxindex);
		for (;;) {
			if (newindex >= maxindex) return true;
			if (IsEntryVailed(*m_ppxa, newindex) == true) {
				m_index = newindex;
				return false;
			}
			newindex++;
		}
	}

public:
	int EnumMode;
#if 0
	ref class CGetDateTime
	{
	private:
		DateTime^ m_value;

	public:
		CGetDateTime(DateTime^ value)
		{
			m_value = value;
		}
		virtual String^ ToString() override
		{
			return m_value->ToString();
		}
		String^ toString()
		{
			return m_value->ToString();
		}
		__int64 getTime()
		{
			return (m_value->ToUniversalTime().Ticks - 621355968000000000i64) / 10000;
		}
		property DateTime^ datetime
		{
			DateTime^ get() {
				return m_value;
			}
		}
	};
#endif
	CEntry(PPXAPPINFOW** ppxa, int index)
	{
		int value;
		m_ppxa = ppxa;
		EnumMode = 0;
		m_enum = 0;

		if (index == -1) { // カーソル位置を初期化
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_CSRINDEX, &value);
			m_index = value;
		}
		else {
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_DIRTTOTAL, &value);
			m_index = (index < value) ? index : ((value > 0) ? value - 1 : 0);
		}
	};

	virtual IEnumerator^ GetEnumerator()
	{
		return gcnew CEntryEnum(m_ppxa, EnumMode);
	}

	property CEntry^ AllMark
	{
		CEntry^ get() {
			CEntry^ entry = gcnew CEntry(m_ppxa, m_index);
			entry->EnumMode = 1;
			return entry;
		}
	}

	property CEntry^ AllEntry
	{
		CEntry^ get() {
			CEntry^ entry = gcnew CEntry(m_ppxa, m_index);
			entry->EnumMode = 2;
			return entry;
		}
	}

	property int length
	{
		int get() {
			int value;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_DIRTTOTAL, &value);
			return value;
		}
	}
	property int Count { int get() { return length; } }

	property int index
	{
		int get() {
			int Value;

			Value = m_index;
			if ((*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRY_HS_GETINDEX, &Value) == 0) {
				return m_index;
			}
			return Value;
		}
		void set(int index) {
			if (index >= 0) {
				int maxindex;

				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_DIRTTOTAL, &maxindex);
				if (index < maxindex) m_index = index;
			}
		}
	}

	property int Index {
		int get() { return index; }
		void set(int num) { index = num; }
	}

	CEntry^ Item(int nIndex)
	{
		return gcnew CEntry(m_ppxa, nIndex);
	}

	CEntry^ item(int nIndex)
	{
		return gcnew CEntry(m_ppxa, nIndex);
	}

	CEntry^ Item()
	{
		return gcnew CEntry(m_ppxa, m_index);
	}

	CEntry^ item()
	{
		return gcnew CEntry(m_ppxa, m_index);
	}

	property String^ Name
	{
		String^ get() {
			PPXUPTR_INDEX_UPATHW tmp;

			tmp.index = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYNAME, &tmp);
			return marshal_as<String^>(tmp.path);
		}
	}

	property String^ ShortName
	{
		String^ get() {
			PPXUPTR_INDEX_UPATHW tmp;

			tmp.index = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYANAME, &tmp);
			if (tmp.path[0] == '\0') {
				tmp.index = m_index;
				(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYNAME, tmp.path);
			}

			return marshal_as<String^>(tmp.path);
		}
	}

	property int Attributes
	{
		int get() {
			int value = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYATTR, &value);
			return value;
		}
	}

	property __int64 Size
	{
		__int64 get() {
			__int64 value = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMSIZE, &value);
			return value;
		}
	}
#if 0
	property CGetDateTime^ DateCreated
	{
		CGetDateTime^ get() {
			__int64 ftime = m_index;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYCTIME, &ftime);
			return gcnew CGetDateTime(DateTime::FromFileTime(ftime));
		}
	}

	property CGetDateTime^ DateLastAccessed
	{
		CGetDateTime^ get() {
			__int64 ftime = m_index;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYATIME, &ftime);
			return gcnew CGetDateTime(DateTime::FromFileTime(ftime));
		}
	}
	property CGetDateTime^ DateLastModified
	{
		CGetDateTime^ get() {
			__int64 ftime = m_index;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMTIME, &ftime);
			return gcnew CGetDateTime(DateTime::FromFileTime(ftime));
		}
	}
#else
#if 0
	property __int64 DateLastModified
	{
		Object^ get() {
			__int64 ftime = m_index;
			V8ScriptEngine^ engine;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMTIME, &ftime);
			engine = gcnew V8ScriptEngine;
			return engine->Evaluate("", "new Date(" + ((ftime - 116444736000000000) / 10000) + ");");
		}
	}
#else
	property __int64 DateCreated
	{
		__int64 get() {
			__int64 ftime = m_index;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYCTIME, &ftime);
			return (ftime - 116444736000000000) / 10000;
		}
	}

	property __int64 DateLastAccessed
	{
		__int64 get() {
			__int64 ftime = m_index;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYATIME, &ftime);
			return (ftime - 116444736000000000) / 10000;
		}
	}
	property __int64 DateLastModified
	{
		__int64 get() {
			__int64 ftime = m_index;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMTIME, &ftime);
			return (ftime - 116444736000000000) / 10000;
		}
	}
#endif
#endif
	property int Mark
	{
		int get() {
			int Value;

			Value = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARK, &Value);
			return Value;
		}
		void set(int value) {
			int nums[2];

			nums[0] = value;
			nums[1] = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSETMARK, nums);
			m_index = nums[1]; // 加工済みの値を回収して高速化
		}
	}

	property int State
	{
		int get() {
			int Value;

			Value = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSTATE, &Value);
			return Value & 0x1f;
		}
		void set(int value) {
			int nums[2];

			nums[0] = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSTATE, nums);

			nums[0] = (nums[0] & 0xffe0) | value;
			nums[1] = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSETSTATE, nums);
			m_index = nums[1]; // 加工済みの値を回収して高速化
		}
	}

	property int Highlight
	{
		int get() {
			int Value;

			Value = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSTATE, &Value);
			return Value >> 5;
		}
		void set(int value) {
			int nums[2];

			nums[0] = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSTATE, nums);

			nums[0] = (nums[0] & 0x1f) | (value << 5);
			nums[1] = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSETSTATE, nums);
			m_index = nums[1]; // 加工済みの値を回収して高速化
		}
	}

	property int ExtColor
	{
		int get() {
			int Value;

			Value = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYEXTCOLOR, &Value);
			return Value;
		}
		void set(int value) {
			int nums[2];

			nums[0] = value;
			nums[1] = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSETEXTCOLOR, nums);
			m_index = nums[1]; // 加工済みの値を回収して高速化
		}
	}

	property String^ Comment
	{
		String^ get() {
			WCHAR bufw[CMDLINESIZE];

			*((int*)bufw) = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYCOMMENT, bufw);
			return marshal_as<String^>(bufw);
		}

		void set(String^ str)
		{
			WCHAR bufw[CMDLINESIZE];
			UINT len;

			if (str == nullptr) str = "";
			*((int*)bufw) = m_index;
			len = str->Length;
			if (len > (CMDLINESIZE - 3)) len = CMDLINESIZE - 3;
			if (len != 0) {
				marshal_context ctx;

				memcpy(bufw + 2, ctx.marshal_as<const wchar_t*>(str), sizeof(WCHAR) * len);
			}
			bufw[len + 2] = '\0';
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYSETCOMMENT, bufw);
			return;
		}
	}

	property String^ Information
	{
		String^ get() {
			PPXUPTR_ENTRYINFOW entryinfo;
			String^ str;

			entryinfo.index = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYINFO, &entryinfo);
			if ((DWORD_PTR)entryinfo.result < 0x10000) return "";
			str = marshal_as<String^>(entryinfo.result);
			HeapFree(GetProcessHeap(), 0, entryinfo.result);
			return str;
		}
	}

	String^ GetComment(int id)
	{
		ENTRYEXTDATASTRUCT eeds;
		WCHAR bufw[CMDLINESIZE];

		eeds.entry = m_index;
		eeds.id = (id > 100) ? id : (DFC_COMMENTEX_MAX - (id - 1));
		eeds.size = sizeof(bufw);
		eeds.data = (BYTE*)bufw;
		if ((*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYEXTDATA_GETDATA, &eeds) == 0) {
			bufw[0] = '\0';
		}
		return marshal_as<String^>(bufw);
	}

	void SetComment(int id, String^ str)
	{
		ENTRYEXTDATASTRUCT eeds;
		marshal_context ctx;

		if (str == nullptr) str = "";
		eeds.entry = m_index;
		eeds.id = (id > 100) ? id : (DFC_COMMENTEX_MAX - (id - 1));
		eeds.size = (str->Length + 1) * sizeof(WCHAR);
		if (eeds.size == sizeof(WCHAR)) { // 空文字列
			eeds.data = (BYTE*)L"";
		}
		else {
			eeds.data = (BYTE*)ctx.marshal_as<const wchar_t*>(str);
		}
		(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYEXTDATA_SETDATA, &eeds);
		return;
	}

	property int Hide
	{
		int get()
		{
			int value = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYHIDEENTRY, &value);
			return m_index;
		}
	}

	property int FirstMark
	{
		int get()
		{
			int result;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKFIRST_HS, &result);
			if (result == -1) {
				result = 0;
			}
			else {
				m_index = result;
				result = 1;
			}
			return result;
		}
	}

	property int NextMark
	{
		int get()
		{
			int result;

			result = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKNEXT_HS, &result);
			if (result == -1) {
				result = 0;
			}
			else {
				m_index = result;
				result = 1;
			}
			return result;
		}
	}

	property int PrevMark
	{
		int get()
		{
			int result;

			result = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKPREV_HS, &result);
			if (result == -1) {
				result = 0;
			}
			else {
				m_index = result;
				result = 1;
			}
			return result;
		}
	}

	property int LastMark
	{
		int get()
		{
			int result;

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKLAST_HS, &result);
			if (result == -1) {
				result = 0;
			}
			else {
				m_index = result;
				result = 1;
			}
			return result;
		}
	}

	void Reset()
	{
		m_index = -1;
		m_enum = 0;
	}

	bool atEnd()
	{
		int result;

		if (m_enum == 0) { // 初めて
			m_enum = 1;
			if (EnumMode == 2) return CheckNextEntryAt(0);

			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKFIRST_HS, &result);
			if (result == -1) {
				if (EnumMode == 0) {
					(*m_ppxa)->Function(*m_ppxa, PPXCMDID_CSRINDEX, &result);
					m_index = result;
				}
				else {
					return true;
				}
			}
			else {
				m_index = result;
			}
			if (IsEntryVailed(*m_ppxa, m_index) == true) return false;
		}

		// ２回目以降
		if (EnumMode == 2) return CheckNextEntryAt(m_index + 1);
		for (;;) {
			result = m_index;
			(*m_ppxa)->Function(*m_ppxa, PPXCMDID_ENTRYMARKNEXT_HS, &result);
			if ((result == -1) || (result == m_index)) return true;
			m_index = result;
			if (IsEntryVailed(*m_ppxa, result) == true) return false;
		}
	}

	void moveNext()
	{
	};
};

public ref class CEnumerator
{
public:
	CEnumerator(Object^ object)
	{
		m_Object = static_cast<IEnumerable^>(object)->GetEnumerator();
	}

	Object^ Item()
	{
		return m_Object->Current;
	}

	Object^ item()
	{
		return m_Object->Current;
	}

	void Reset()
	{
		m_Object->Reset();
	}

	bool atEnd()
	{
		return !m_Object->MoveNext();
	}

	void moveNext()
	{
	}

	void moveFirst()
	{
		m_Object->Reset();
	}

private:
	IEnumerator^ m_Object;
};

public ref class CResult
{
public:
	CResult()
	{
		UseResult = false;
	}

	property String^ result
	{
		String^ get()
		{
			if (UseResult == false) return "";
			return resultvalue;
		}
		void set(String^ str)
		{
			resultvalue = str;
			UseResult = true;
		}

	}
	bool UseResult;
	String^ resultvalue;
};

public ref class CPPxObjects
{
private:
	InstanceValueStruct* m_info;
	CInstance^ m_instance;
	V8ScriptEngine^ m_engine;
	PPXAPPINFOW* m_ppxa;
	CResult^ m_value;
	ERRORCODE m_extract_result;
	Collections::Generic::List<CEventInstance^>^ m_events;

public:
	CPPxObjects(InstanceValueStruct* info, CInstance^ instance, V8ScriptEngine^ engine, CResult^ value) {
		m_info = info;
		m_instance = instance;
		m_engine = engine;
		m_ppxa = info->ppxa;
		m_value = value;
		m_extract_result = NO_ERROR;
	}

	~CPPxObjects()
	{
		this->!CPPxObjects();
	}

	!CPPxObjects()
	{
		//	report("!CPPxObjects\r\n");
			// 空の定義をすることで、m_event のファイナライザを起動させる
	}

	void UpdatePPxInfo(void)
	{
		m_ppxa = m_info->ppxa;
	}

	property CArguments^ Arguments
	{
		CArguments^ get() {
			return gcnew CArguments(&m_info->ppxa, &m_info->pxc);
		}
	}

	String^ Argument(int nIndex)
	{
		return GetArgumentsItem(m_info->ppxa, m_info->pxc, nIndex);
	}

	property CEntry^ Entry
	{
		CEntry^ get() {
			return gcnew CEntry(&m_info->ppxa, -1);
		}
	}

	property CEntry^ entry
	{
		CEntry^ get() {
			return gcnew CEntry(&m_info->ppxa, -1);
		}
	}

	property CPane^ Pane
	{
		CPane^ get() {
			return gcnew CPane(&m_info->ppxa, -1);
		}
	}

	property CPane^ pane
	{
		CPane^ get() {
			return gcnew CPane(&m_info->ppxa, -1);
		}
	}

	property int ReentryCount
	{
		int get() {
			return m_info->stay.entry;
		}
	}

	void Include(String^ filename)
	{
		DocumentInfo codedocs(filename);
		m_engine->Execute(codedocs, GetScriptText(filename, 1));
	}

	void DisconnectObject(Object^ object)
	{
		if (m_events == nullptr) return; // 登録無し
		int index, maxindex;
		IDispatch* pDispatch;

		maxindex = m_events->Count;
		for (index = 0; index < maxindex; index++) {
			pDispatch = static_cast<IDispatch*>(Marshal::GetIDispatchForObject(object).ToPointer());

			if (m_events[index]->SourceObject->Equals(object)) {
				m_events[index]->Disconnect();
				m_events->RemoveAt(index);
			}
		}
	}

	void ConnectObject(Object^ object, String^ prefix)
	{
		if (m_info->ModuleMode) throw gcnew Exception("Can't use event in moudlue mode");

		IDispatch* pDispatch = static_cast<IDispatch*>(Marshal::GetIDispatchForObject(object).ToPointer());

		CEventInstance^ event = gcnew CEventInstance(m_engine);

		::IConnectionPointContainer* pContainer;
		if (SUCCEEDED(pDispatch->QueryInterface(IID_IConnectionPointContainer,
			reinterpret_cast<void**>(&pContainer)))) {
			::IEnumConnectionPoints* pEnumPoints;

			if (SUCCEEDED(pContainer->EnumConnectionPoints(&pEnumPoints))) {
				::IConnectionPoint* pPoint;

				if (SUCCEEDED(pEnumPoints->Next(1, &pPoint, 0))) {
					IID eventID;
					::ITypeInfo* pDispType, * pEventType;
					::ITypeLib* pTLib;
					UINT libindex;

					pPoint->GetConnectionInterface(&eventID);
					pDispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pDispType);
					pDispType->GetContainingTypeLib(&pTLib, &libindex);
					pDispType->Release();
					pTLib->GetTypeInfoOfGuid(eventID, &pEventType);
					pTLib->Release();

					if (SUCCEEDED(event->Connect(object, pPoint, pEventType, prefix))) {
						if (m_events == nullptr) {
							m_events = gcnew Collections::Generic::List<CEventInstance^>();
						}
						m_events->Add(event);
					}
				}
				pEnumPoints->Release();
			}
			pContainer->Release();
		}
		pDispatch->Release();
	}

	Object^ CreateObject(String^ objectname)
	{
		return Activator::CreateInstance(Type::GetTypeFromProgID(objectname));
	}

	Object^ CreateObject(String^ objectname, String^ prefix)
	{
		Object^ obj = Activator::CreateInstance(Type::GetTypeFromProgID(objectname));
		ConnectObject(obj, prefix);
		return obj;
	}

	Object^ GetObject(String^ filename, String^ objectname)
	{
		return Activator::CreateInstanceFrom(filename, objectname);
	}

	Object^ GetObject(String^ filename, String^ objectname, String^ prefix)
	{
		Object^ obj = Activator::CreateInstanceFrom(filename, objectname);
		ConnectObject(obj, prefix);
		return obj;
	}

	void AddReference(String^ objectname, String^ filename, String^ assembly)
	{
		m_engine->AddHostObject(objectname, gcnew HostTypeCollection(filename, assembly));
	}

	void AddComObject(String^ objectname, String^ objectID)
	{
		m_engine->AddCOMObject(objectname, objectID);
	}

	Object^ Enum(Object^ objectname)
	{
		return gcnew CEnumerator(objectname);
	}

	Object^ Enumerator(Object^ objectname)
	{
		return gcnew CEnumerator(objectname);
	}

	void Echo(String^ str)
	{
		::PopupMessage(m_ppxa, (str != nullptr) ? str : "");
	}

	void Echo(Object^ value)
	{
		::PopupMessage(m_ppxa, (value != nullptr) ? value->ToString() : "");
	}

	void echo(Object^ value)
	{
		::PopupMessage(m_ppxa, (value != nullptr) ? value->ToString() : "");
	}

	void echo(... array<Object^>^ args)
	{
		Echo(args);
	}

	void Echo(... array<Object^>^ args)
	{
		System::Text::StringBuilder^ str = gcnew System::Text::StringBuilder();
		bool space = false;
		for each (Object ^ arg in args) {
			if (arg != nullptr) {
				if (space) {
					str->Append(" ");
				}
				else {
					space = true;
				}
				str->Append(arg);
			}
		}
		::PopupMessage(m_ppxa, str->ToString());
	}

	void Sleep(int n)
	{
		::Sleep(n);
	}

	Object^ option(String^ name)
	{
		return nullptr;
	}

	void Quit()
	{
		throw gcnew Exception("quitdone");
	}

	void quit()
	{
		Quit();
	}

	void Quit(int mode)
	{
		throw gcnew Exception(mode == 1 ? "quitdone" : "quitstop");
	}

	void quit(int mode)
	{
		Quit(mode);
	}

	property String^ ScriptName
	{
		String^ get() {
			return marshal_as<String^>(m_info->pxc->param);
		}
	}

	property String^ ScriptFullName
	{
		String^ get() {
			return marshal_as<String^>(m_info->pxc->param);
		}
	}

	property int CodeType
	{
		int get() {
			return static_cast<int>(m_ppxa->Function(m_ppxa, PPXCMDID_CHARTYPE, NULL));
		}
	}

	property int ModuleVersion
	{
		int get() {
			return SCRIPTMODULEVER;
		}
	}

	property int PPxVersion
	{
		int get()
		{
			return static_cast<int>(m_ppxa->Function(m_ppxa, PPXCMDID_VERSION, NULL));
		}
	}

	property String^ ScriptEngineName
	{
		String^ get()
		{
#if 0
			V8ScriptEngine^ engine;

			engine = gcnew V8ScriptEngine;
			return engine->exce
				IScriptEngineException::EngineName;
			V8::IScriptEngineException::EngineName;
#endif
			return "ClearScriptV8";
		}
	}

	property String^ ScriptEngineVersion
	{
		String^ get()
		{
#if 0
			V8ScriptEngine^ engine;

			engine = gcnew V8ScriptEngine;
			return engine->exce
				IScriptEngineException::EngineName;
			V8::IScriptEngineException::EngineName;
#endif
			return "7.2.2";
		}
	}

	property Object^ result
	{
		Object^ get() {
			return m_value->result;
		}
		void set(Object^ value) {
			if ((value == nullptr) || (value->GetType() == Microsoft::ClearScript::Undefined::typeid)) {
				m_value->result = "";
			}
			else if (value->GetType() == bool::typeid) {
				m_value->result = safe_cast<bool>(value) ? L"-1" : L"0";
			}
			else {
				m_value->result = value->ToString();
			}
			return;
		}
	}

	property Object^ Result
	{
		Object^ get() { return result; }
		void set(Object^ str) { result = str; }
	}

	property int StayMode
	{
		int get() {
			return m_info->stay.mode;
		}
		void set(int index); // CInstance の詳細が無いので後で実体を記載
	}

	int Execute(String^ str)
	{
		marshal_context ctx;
		int result;

		if (str == nullptr) return NO_ERROR;
		result = static_cast<int>(m_ppxa->Function(m_ppxa, PPXCMDID_EXECUTE, (void*)(ctx.marshal_as<const wchar_t*>(str))));
		if (result <= 1) result ^= 1;
		return result;
	}

	int execute(String^ str)
	{
		return Execute(str);
	}

	String^ Extract(String^ str)
	{
		BSTR b_result;
		String^ s_result;
		marshal_context ctx;
		const wchar_t* wstr;

		if (str == nullptr) return "";
		wstr = ctx.marshal_as<const wchar_t*>(str);
		b_result = (BSTR)m_ppxa->Function(m_ppxa, PPXCMDID_LONG_EXTRACT_E, (void*)wstr);
		if (b_result == NULL) {
			b_result = (BSTR)m_ppxa->Function(m_ppxa, PPXCMDID_LONG_EXTRACT, (void*)wstr);
			if (b_result == NULL) return "";
		}

		if (((DWORD_PTR)b_result & ~(DWORD_PTR)0xffff) == 0) {
			m_extract_result = (ERRORCODE)(DWORD_PTR)b_result;
			return "";
		}
		m_extract_result = NO_ERROR;

		s_result = marshal_as<String^>(b_result);
		::SysFreeString(b_result);
		return s_result;
	}

	String^ extract(String^ str)
	{
		return Extract(str);
	}

	int Extract()
	{
		return m_extract_result;
	}

	void SetPopLineMessage(String^ str)
	{
		marshal_context ctx;
		if (str == nullptr) return;
		m_ppxa->Function(m_ppxa, PPXCMDID_SETPOPLINE, (void*)ctx.marshal_as<const wchar_t*>(str));
		return;
	}

	void SetPopLineMessage(Object^ value)
	{
		marshal_context ctx;
		if (value == nullptr) return;
		m_ppxa->Function(m_ppxa, PPXCMDID_SETPOPLINE, (void*)ctx.marshal_as<const wchar_t*>(value->ToString()));
		return;
	}

	void linemessage(String^ str)
	{
		marshal_context ctx;
		if (str == nullptr) return;
		m_ppxa->Function(m_ppxa, PPXCMDID_SETPOPLINE, (void*)ctx.marshal_as<const wchar_t*>(str));
		return;
	}

	void linemessage(Object^ value)
	{
		marshal_context ctx;
		if (value == nullptr) return;
		m_ppxa->Function(m_ppxa, PPXCMDID_SETPOPLINE, (void*)ctx.marshal_as<const wchar_t*>(value->ToString()));
		return;
	}

	void report(Object^ value)
	{
		marshal_context ctx;
		if (value == nullptr) return;
		m_ppxa->Function(m_ppxa, PPXCMDID_REPORTTEXT, (void*)ctx.marshal_as<const wchar_t*>(value->ToString()));
		return;
	}

	void log(Object^ value)
	{
		marshal_context ctx;
		if (value == nullptr) return;
		m_ppxa->Function(m_ppxa, PPXCMDID_DEBUGLOG, (void*)ctx.marshal_as<const wchar_t*>(value->ToString()));
		return;
	}

	String^ GetFileInformation(String^ str, int mode)
	{
		WCHAR bufw[CMDLINESIZE];
		marshal_context ctx;

		*(int*)bufw = mode;
		bufw[2] = '\0';
		wcscpy(bufw + 2, ctx.marshal_as<const wchar_t*>(str));
		m_ppxa->Function(m_ppxa, PPXCMDID_GETFILEINFO, bufw);
		return marshal_as<String^>(bufw + 2);
	}

	String^ GetFileInformation(String^ str)
	{
		return GetFileInformation(str, 0);
	}

	property String^ Clipboard
	{
		String^ get() {
			if (TryOpenClipboard(m_ppxa->hWnd) != FALSE) {
				HANDLE clipdata;
				WCHAR* src;

				clipdata = GetClipboardData(CF_UNICODETEXT);
				if (clipdata != NULL) {
					String^ str;

					src = (WCHAR*)GlobalLock(clipdata);
					str = marshal_as<String^>(src);
					GlobalUnlock(clipdata);
					CloseClipboard();
					return str;
				}
				CloseClipboard();
			}
			return "";
		}

		void set(String^ str) {
			marshal_context ctx;
			HGLOBAL hm;
			size_t len;

			if (str == nullptr) str = "";
			len = str->Length * sizeof(WCHAR) + sizeof(WCHAR); // NIL分を追加する
			hm = GlobalAlloc(GMEM_MOVEABLE, len);
			if (hm == NULL) return;
			memcpy(GlobalLock(hm), ctx.marshal_as<const wchar_t*>(str), len);
			GlobalUnlock(hm);
			if (TryOpenClipboard(m_ppxa->hWnd)) {
				EmptyClipboard();
				SetClipboardData(CF_UNICODETEXT, hm);
				CloseClipboard();
			}
			return;
		}
	}

	property String^ WindowIDName
	{
		String^ get() {
			WCHAR bufW[REGEXTIDSIZE];

			if ((m_ppxa->Function(m_ppxa, PPXCMDID_GETSUBID, &bufW) == 0) || ((UTCHAR)bufW[0] > 'z')) {
				if (m_ppxa->RegID[1] != '_') {
					return marshal_as<String^>(m_ppxa->RegID);
				}
				bufW[0] = m_ppxa->RegID[0];
				strcpyW(bufW + 1, m_ppxa->RegID + 2);
			}
			return marshal_as<String^>(bufW);
		}
		void set(String^) {
		}
	}

	property int WindowDirection
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_WINDOWDIR, &value);
			return value;
		}
	}

	property int EntryDisplayTop
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRDINDEX, &value);
			return value;
		}
		void set(int index)
		{
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSETDINDEX, &index);
		}
	}

	property int EntryDisplayX
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRDISPW, &value);
			return value;
		}
	}

	property int EntryDisplayY
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRDISPH, &value);
			return value;
		}
	}

	int GetComboItemCount() {
		DWORD_PTR value;

		value = 0;
		m_ppxa->Function(m_ppxa, PPXCMDID_COMBOIDCOUNT, &value);
		return static_cast<int>(value);
	}

	int GetComboItemCount(int mode) {
		DWORD_PTR value;

		value = mode;
		m_ppxa->Function(m_ppxa, PPXCMDID_COMBOIDCOUNT, &value);
		return static_cast<int>(value);
	}

	property String^ ComboIDName
	{
		String^ get() {
			WCHAR bufw[8];

			m_ppxa->Function(m_ppxa, PPXCMDID_COMBOIDNAME, bufw);
			return marshal_as<String^>(bufw);
		}
	}

	property int SlowMode
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_SLOWMODE, &value);
			return value;
		}
		void set(int mode)
		{
			int value = mode;
			m_ppxa->Function(m_ppxa, PPXCMDID_SETSLOWMODE, &value);
		}
	}

	property int SyncView
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_SYNCVIEW, &value);
			return value;
		}
		void set(int mode)
		{
			int value = mode;
			m_ppxa->Function(m_ppxa, PPXCMDID_SETSYNCVIEW, &value);
		}
	}

	property String^ DriveVolumeLabel
	{
		String^ get() {
			WCHAR bufw[CMDLINESIZE];

			bufw[0] = '\0';
			m_ppxa->Function(m_ppxa, PPXCMDID_DRIVELABEL, bufw);
			return marshal_as<String^>(bufw);
		}
	}

	property __int64 DriveTotalSize
	{
		__int64 get() {
			__int64 size;

			m_ppxa->Function(m_ppxa, PPXCMDID_DRIVETOTALSIZE, &size);
			return size;
		}
	}

	property __int64 DriveFreeSize
	{
		__int64 get() {
			__int64  size;

			m_ppxa->Function(m_ppxa, PPXCMDID_DRIVEFREE, &size);
			return size;
		}
	}

	UINT LoadCount(int type)
	{
		int value = type;
		m_ppxa->Function(m_ppxa, PPXCMDID_DIRLOADCOUNT, &value);
		return value;
	}

	property int DirectoryType
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_DIRTYPE, &value);
			return value;
		}
	}

	property int EntryAllCount
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_DIRTOTAL, &value);
			return value;
		}
	}

	property int EntryDisplayCount
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_DIRTTOTAL, &value);
			return value;
		}
	}

	property int EntryMarkCount
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_DIRMARKS, &value);
			return value;
		}
	}

	property __int64 EntryMarkSize
	{
		__int64 get() {
			__int64  size;

			m_ppxa->Function(m_ppxa, PPXCMDID_DIRMARKSIZE, &size);
			return size;
		}
	}

	property int EntryDisplayDirectories
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_DIRTTOTALDIR, &value);
			return value;
		}
	}

	property int EntryDisplayFiles
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_DIRTTOTALFILE, &value);
			return value;
		}
	}

	void EntryInsert(int index, String^ str)
	{
		DWORD_PTR dptrs[2];
		marshal_context ctx;

		dptrs[0] = static_cast<DWORD_PTR>(index);
		dptrs[1] = (DWORD_PTR)(void*)ctx.marshal_as<const wchar_t*>(str);
		m_ppxa->Function(m_ppxa, PPXCMDID_ENTRYINSERTMSG, dptrs);
		return;
	}

	property int EntryIndex
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRINDEX, &value);
			return value;
		}
		void set(int i)
		{
			int value = i;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSETINDEX, &value);
		}
	}

	property int EntryMark
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRMARK, &value);
			return value;
		}
		void set(int i)
		{
			int value = i;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSETMARK, &value);
		}
	}

	property String^ EntryName
	{
		String^ get() {
			WCHAR bufw[CMDLINESIZE];
			PPXCMDENUMSTRUCTW info;

			info.buffer = bufw;
			m_ppxa->Function(m_ppxa, 'R', &info);
			return marshal_as<String^>(bufw);
		}
	}

	property int EntryAttributes
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRATTR, &value);
			return value;
		}
	}

	property __int64 EntrySize
	{
		__int64 get() {
			__int64 value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRMSIZE, &value);
			return value;
		}
	}

	property int EntryState
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSTATE, &value);
			return value;
		}
		void set(int i)
		{
			int tmpvalue;
			int value;

			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSTATE, &tmpvalue);
			value = (tmpvalue & 0xffe0) | i;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSETSTATE, &value);
		}
	}

	property int EntryExtColor
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSREXTCOLOR, &value);
			return value;
		}
		void set(int i)
		{
			int value = i;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSETEXTCOLOR, &value);
		}
	}

	property int EntryHighlight
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSTATE, &value);
			return value >> 5;
		}
		void set(int i)
		{
			int tmpvalue;
			int value;

			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSTATE, &tmpvalue);
			value = (tmpvalue & 0x1f) | (i << 5);
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSETSTATE, &value);
		}
	}

	property String^ EntryComment
	{
		String^ get() {
			WCHAR bufw[CMDLINESIZE];

			m_ppxa->Function(m_ppxa, PPXCMDID_CSRCOMMENT, bufw);
			return marshal_as<String^>(bufw);
		}

		void set(String^ str)
		{
			marshal_context ctx;
			if (str == nullptr) str = "";
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRSETCOMMENT, (void*)ctx.marshal_as<const wchar_t*>(str));
			return;
		}
	}

	property int EntryFirstMark
	{
		int get() {
			int value;
			if (PPXA_NO_ERROR == m_ppxa->Function(m_ppxa, PPXCMDID_CSRMARKFIRST, &value)) {
				return value;
			}
			else {
				throw gcnew Exception("not impliment");
			}
		}
	}

	property int EntryNextMark
	{
		int get() {
			int value;
			if (PPXA_NO_ERROR == m_ppxa->Function(m_ppxa, PPXCMDID_CSRMARKNEXT, &value)) {
				return value;
			}
			else {
				throw gcnew Exception("not impliment");
			}
		}
	}

	property int EntryPrevMark
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRMARKPREV, &value);
			return value;
		}
	}

	property int EntryLastMark
	{
		int get() {
			int value;
			m_ppxa->Function(m_ppxa, PPXCMDID_CSRMARKLAST, &value);
			return value;
		}
	}

	property int PointType
	{
		int get() {
			int var[3];

			m_ppxa->Function(m_ppxa, PPXCMDID_POINTINFO, var);
			return var[0];
		}
	}

	property int PointIndex
	{
		int get() {
			int var[3];

			m_ppxa->Function(m_ppxa, PPXCMDID_POINTINFO, var);
			return var[1];
		}
	}

	property Object^ value[String^]
	{
		Object ^ get(String ^ key)
		{
			if (gc_value == nullptr) return nullptr;
			return gc_value[key];
		}

		void set(String ^ key, Object ^ value)
		{
			if (gc_value == nullptr) {
				g_value = gcnew Collections::Concurrent::ConcurrentDictionary<String^, Object^>();
			}
			gc_value[key] = value;
		}
	}

		Object^ getValue(String^ key)
	{
		if (gc_value == nullptr) return nullptr;
		return gc_value[key];
	}

	void setValue(String^ key, Object^ value)
	{
		if (gc_value == nullptr) {
			g_value = gcnew Collections::Concurrent::ConcurrentDictionary<String^, Object^>();
		}
		gc_value[key] = value;
	}

	String^ getProcessValue(String^ key)
	{
		marshal_context ctx;
		void* uptr[2];
		WCHAR buf[CMDLINESIZE];
		String^ str;

		uptr[0] = (void*)ctx.marshal_as<const wchar_t*>(key);
		uptr[1] = buf;
		if (0 == m_ppxa->Function(m_ppxa, PPXCMDID_GETPROCVARIABLEDATA, uptr)) {
			buf[0] = '\0';
		}
		str = marshal_as<String^>((WCHAR*)uptr[1]);
		if (uptr[1] != buf) HeapFree(GetProcessHeap(), 0, uptr[1]);
		return str;
	}

	void setProcessValue(String^ key, String^ value)
	{
		marshal_context ctx_key, ctx_value;
		void* uptr[2];

		uptr[0] = (void*)ctx_key.marshal_as<const wchar_t*>(key);
		uptr[1] = (void*)ctx_value.marshal_as<const wchar_t*>(value);
		m_ppxa->Function(m_ppxa, PPXCMDID_SETPROCVARIABLEDATA, uptr);
	}

	String^ getIValue(String^ key)
	{
		marshal_context ctx;
		void* uptr[2];
		WCHAR buf[CMDLINESIZE];
		String^ str;

		uptr[0] = (void*)ctx.marshal_as<const wchar_t*>(key);
		uptr[1] = buf;
		if (0 == m_ppxa->Function(m_ppxa, PPXCMDID_GETWNDVARIABLEDATA, uptr)) {
			buf[0] = '\0';
		}
		str = marshal_as<String^>((WCHAR*)uptr[1]);
		if (uptr[1] != buf) HeapFree(GetProcessHeap(), 0, uptr[1]);
		return str;
	}

	void setIValue(String^ key, String^ value)
	{
		marshal_context ctx_key;
		marshal_context ctx_value;
		void* uptr[2];

		uptr[0] = (void*)ctx_key.marshal_as<const wchar_t*>(key);
		uptr[1] = (void*)ctx_value.marshal_as<const wchar_t*>(value);
		m_ppxa->Function(m_ppxa, PPXCMDID_SETWNDVARIABLEDATA, uptr);
	}
};


VC_DLL_EXPORTS int WINAPI ModuleEntry(PPXAPPINFOW* ppxa, DWORD cmdID, PPXMODULEPARAM pxs)
{
	if (cmdID == PPXMEVENT_COMMAND) {
		switch (pxs.command->commandhash) {
		case 0x800012d3:
			if (!strcmpW(pxs.command->commandname, L"JS")) {
				return RunScript(ppxa, pxs.command, 0);
			}
			break;

		case 0x8004b4f8:
			if (!::wcscmp(pxs.command->commandname, L"JS8")) {
				return RunScript(ppxa, pxs.command, 0);
			}
			break;

		case 0xc34c9454:
			if (!strcmpW(pxs.command->commandname, L"SCRIPT")) {
				return RunScript(ppxa, pxs.command, 1);
			}
			break;

		case 0xd3251538:
			if (!strcmpW(pxs.command->commandname, L"SCRIPT8")) {
				return RunScript(ppxa, pxs.command, 1);
			}
			break;

		case 0xd3251571:
			if (!strcmpW(pxs.command->commandname, L"SCRIPTA")) {
				return RunScript(ppxa, pxs.command, 2);
			}
			break;

		case 0xdff577dd:
			if (!strcmpW(pxs.command->commandname, L"DISCARDMODULECACHE")) {
				V8ScriptEngine^ engine;

				engine = gcnew V8ScriptEngine();
				engine->DocumentSettings->Loader->DiscardCachedDocuments();
				return PPXMRESULT_DONE;
			}
			break;
		}
		return PPXMRESULT_SKIP;
	}
	if (cmdID == PPXMEVENT_FUNCTION) {
		switch (pxs.command->commandhash) {
		case 0x800012d3:
			if (!strcmpW(pxs.command->commandname, L"JS")) {
				return RunScript(ppxa, pxs.command, 0);
			}
			break;

		case 0x8004b4f8:
			if (!::wcscmp(pxs.command->commandname, L"JS8")) {
				return RunScript(ppxa, pxs.command, 0);
			}
			break;

		case 0xc34c9454:
			if (!strcmpW(pxs.command->commandname, L"SCRIPT")) {
				return RunScript(ppxa, pxs.command, 1);
			}
			break;

		case 0xd3251538:
			if (!strcmpW(pxs.command->commandname, L"SCRIPT8")) {
				return RunScript(ppxa, pxs.command, 1);
			}
			break;

		case 0xd3251571:
			if (!strcmpW(pxs.command->commandname, L"SCRIPTA")) {
				return RunScript(ppxa, pxs.command, 2);
			}
			break;
		}
		return PPXMRESULT_SKIP;
	}

	if ( cmdID == PPXMEVENT_CLOSETHREAD ){ // 1.97+1 から有効
		FreeStayInstance();
		return PPXMRESULT_DONE;
	}

	if (cmdID == PPXM_INFORMATION) {
		if (pxs.info->infotype == 0) {
			pxs.info->typeflags = PPMTYPEFLAGS(PPXM_INFORMATION) | PPMTYPEFLAGS(PPXMEVENT_COMMAND) | PPMTYPEFLAGS(PPXMEVENT_FUNCTION);
			wcscpy(pxs.info->copyright, L"PPx V8 Script Module R" SCRIPTMODULEVERSTR L"  Copyright (c)TORO");
			return PPXMRESULT_DONE;
		}
	}
	if (cmdID == PPXMEVENT_CLEANUP) {
		if ( PPxVersion > 19700 ) FreeStayInstance();
		if (gc_value != nullptr) delete g_value; // 解放
		return PPXMRESULT_DONE;
	}
	return PPXMRESULT_SKIP;
}

int AfterException(Exception^ e, PPXAPPINFOW* pxa, String^ source)
{
	if (e->Message == "Error: quitdone") {
		return PPXMRESULT_DONE;
	}
	if (e->Message == "Error: quitstop") {
		return PPXMRESULT_STOP;
	}

	String^ mes;
	mes = "PPx V8 実行: " + e->Source + ", " + source + "\n" + e->Message + "\nTarget:\n" + e->TargetSite;

	if (e->InnerException != nullptr) {
		mes += "\nInnerException:\n" + e->InnerException;
	}
	// + "\nStacks:" + e->StackTrace
	::PopupMessage(pxa, mes);
	return PPXMRESULT_STOP;

	// catch (IScriptEngineException^ eu) {
	//	var^ error = e.GetInnerMost<IScriptEngineException>() ? .ErrorDetails;
	//	return error != null ? StringHelpers.CleanupStackTrace(error) : exception.Message;
}

public ref class CInstance
{
public:
	InstanceValueStruct* info;
	V8ScriptEngine^ engine;
	CPPxObjects^ PPxObjects;
	String^ source;

	CInstance(PPXAPPINFOW* ppxa, PPXMCOMMANDSTRUCT* pxc, CResult^ value, int StayMode)
	{
		info = new InstanceValueStruct;
		info->ppxa = ppxa;
		info->pxc = pxc;
		info->ModuleMode = 0;
		info->stay.mode = StayMode;
		info->stay.entry = 0;
		info->stay.threadID = 0;
		info->stay.ppxa = NULL;

		engine = gcnew V8ScriptEngine(V8ScriptEngineFlags::EnableDynamicModuleImports /*| V8ScriptEngineFlags::EnableDateTimeConversion*/);
		engine->DocumentSettings->AccessFlags = DocumentAccessFlags::EnableFileLoading;
		engine->AddHostObject("NETAPI", gcnew HostTypeCollection("mscorlib", "System.Core"));
		PPxObjects = gcnew CPPxObjects(info, this, engine, value);
		engine->AddHostObject("PPx", PPxObjects);
	}

	~CInstance()
	{
		if ( info->stay.ppxa != NULL ){
			HeapFree(GetProcessHeap(), 0, info->stay.ppxa);
		}
		this->!CInstance();
	}
	!CInstance()
	{
		delete info;
	}
};

void CPPxObjects::StayMode::set(int index)
{
	if ( index < 0 ) return;

	if ( index >= ScriptStay_Stay ){
		if ( m_info->stay.threadID == 0 ){
			if ( index == ScriptStay_Stay ){
				index = StayInstanseIDserver++;
				if ( index >= ScriptStay_MaxAutoID ){
					StayInstanseIDserver = ScriptStay_FirstAutoID;
				}
			}
		}
	}else{
		if (m_info->stay.threadID != 0) {
			DropStayInstance(m_instance);
		}
	}
	m_info->stay.mode = index;
}

String^ GetScriptText(String^ param, int file)
{
	if (file == 1) return File::ReadAllText(param);
	Text::Encoding^ def = Text::Encoding::GetEncoding(0);
	return File::ReadAllText(param, def);
}

void GetPPxPath(PPXAPPINFOW* pxa)
{
	WCHAR bufw[CMDLINESIZE];

	bufw[0] = '\0';
	pxa->Function(pxa, PPXCMDID_PPLIBPATH, bufw);
	g_PPxPath = marshal_as<String^>(bufw);
}

#undef SearchPath
String^ InitScriptText(CInstance^ instance, String^% source, int file)
{
	String^ scripttext;

	source = marshal_as<String^>(instance->info->pxc->param);
	if (file) {
		instance->engine->DocumentSettings->SearchPath = Path::GetFullPath(source + "\\..");
		scripttext = GetScriptText(source, file);
	}
	else {
		if (static_cast<String^>(g_PPxPath) == nullptr) {
			GetPPxPath(instance->info->ppxa);
		}
		instance->engine->DocumentSettings->SearchPath = static_cast<String^>(g_PPxPath);
		scripttext = source;
	}
	return scripttext;
}

CInstance^ GetStayInstance(PPXAPPINFOW* ppxa, PPXMCOMMANDSTRUCT* pxc, OldPPxInfoStruct *OldInfo, int StayMode, int file)
{
	DWORD ThreadID = GetCurrentThreadId();
	HWND hWnd = ppxa->hWnd;
	int index, maxindex;
	CInstance^ StayInstance;
	String^ source = marshal_as<String^>(pxc->param);

	System::Threading::Monitor::Enter(gc_StayInstance);
	try {
		maxindex = gc_StayInstance->Count;
		for (index = 0; index < maxindex; index++) {
			StayInstance = gc_StayInstance[index];
			if ((StayInstance->info->stay.threadID == ThreadID) &&
				((StayMode >= ScriptStay_Stay) ?
					(StayInstance->info->stay.mode == StayMode) :
					((StayInstance->info->stay.hWnd == hWnd) &&
					 ((file == 0) || (StayInstance->source == source))) ) ){
				break;
			}
		}
	}
	finally {
		System::Threading::Monitor::Exit(gc_StayInstance);
	}
	if (index >= maxindex) return nullptr;
	OldInfo->ppxa = StayInstance->info->ppxa;
	OldInfo->pxc = StayInstance->info->pxc;
	StayInstance->info->ppxa = ppxa;
	StayInstance->info->pxc = pxc;
	StayInstance->info->stay.entry++;
	StayInstance->PPxObjects->UpdatePPxInfo();
	return StayInstance;
}

DWORD_PTR USECDECL DummyPPxFunc(PPXAPPINFOW* ppxa, DWORD cmdID, void* uptr)
{
	return PPXA_INVALID_FUNCTION;
}

PPXAPPINFOW DummyPPxAppInfo = { DummyPPxFunc, L"", L"", NULL };

PPXMCOMMANDSTRUCT DummyPxc =
#ifndef _WIN64
{ L"", L"", 0, 0, NULL };
#else
{L"", NULL, L"", 0, 0};
#endif

void SleepStayInstance(CInstance^ instance, OldPPxInfoStruct *OldInfo)
{
	// appinfo が使えない状態で使われた時用のダミーを設定する。
	instance->info->ppxa = OldInfo->ppxa;
	instance->info->pxc = OldInfo->pxc;
	instance->info->stay.entry--;
	instance->PPxObjects->UpdatePPxInfo();
}

void ChainStayInstance(CInstance^ instance)
{
	OldPPxInfoStruct OldInfo;

	if ( instance->info->stay.ppxa == NULL ){
		PPXCMDID_NEWAPPINFO_STRUCT pns = {0, 0, NULL, NULL};
		PPXAPPINFOW *ppxa = instance->info->ppxa;

		DWORD_PTR stayppxa = ppxa->Function(ppxa, PPXCMDID_NEWAPPINFO, &pns);
		if ( stayppxa != PPXA_INVALID_FUNCTION ){
			instance->info->stay.ppxa = (PPXAPPINFOW *)stayppxa;
		}
	}
	OldInfo.ppxa = (instance->info->stay.ppxa != NULL) ?
			instance->info->stay.ppxa : &DummyPPxAppInfo;
	OldInfo.pxc = &DummyPxc;
	instance->info->stay.threadID = GetCurrentThreadId();
	instance->info->stay.hWnd = instance->info->ppxa->hWnd;
	PPxVersion = static_cast<long>(instance->info->ppxa->Function(instance->info->ppxa, PPXCMDID_VERSION, NULL));
	SleepStayInstance(instance, &OldInfo);

	if (gc_StayInstance == nullptr) {
		g_StayInstance = gcnew Collections::Generic::List<CInstance^>();
	}
	System::Threading::Monitor::Enter(gc_StayInstance);
	try {
		gc_StayInstance->Add(instance);
	}
	finally {
		System::Threading::Monitor::Exit(gc_StayInstance);
	}
}

// 常駐解除
void DropStayInstance(CInstance^ instance)
{
	int index, maxindex;
	CInstance^ StayInstance;

	System::Threading::Monitor::Enter(gc_StayInstance);
	try {
		maxindex = gc_StayInstance->Count;
		for (index = 0; index < maxindex; index++) {
			StayInstance = gc_StayInstance[index];
			if ( StayInstance->Equals(instance) ){
				gc_StayInstance->RemoveAt(index);
				StayInstance->info->stay.threadID = 0;
				break;
			}
		}
	}
	finally {
		System::Threading::Monitor::Exit(gc_StayInstance);
	}
}

void FreeStayInstance(void)
{
	int index, maxindex;
	CInstance^ StayInstance;
	BOOL changed;
	DWORD ThreadID = GetCurrentThreadId();

	for (;;){
		changed = FALSE;
		System::Threading::Monitor::Enter(gc_StayInstance);
		try {
			maxindex = gc_StayInstance->Count;
			for (index = 0; index < maxindex; index++) {
				StayInstance = gc_StayInstance[index];
				if ( StayInstance->info->stay.threadID == ThreadID ){
					changed = TRUE;
					gc_StayInstance->RemoveAt(index);
					break;
				}
			}
		}
		finally {
			System::Threading::Monitor::Exit(gc_StayInstance);
		}
		if ( !changed ) break;
		try {
			StayInstance->engine->Invoke("ppx_finally");
		}catch (Exception^) { // エラーが起きてもなにもしない
		}
		delete StayInstance;
	}
}

int GetIntNumberW(const WCHAR *line)
{
	int n = 0;

	for ( ;; ){
		WCHAR code;

		code = *line;
		if ( (code < '0') || (code > '9') ) break;
		n = n * 10 + (WCHAR)((BYTE)code - (BYTE)'0');
		line++;
	}
	return n;
}


void CheckOption(PPXMCOMMANDSTRUCT *pxc, int *StayMode, WCHAR *InvokeName)
{
	WCHAR buf[256], *dest;

	pxc->param++;
	for (;;){
		dest = buf;
		for (;;){
			if ( *pxc->param == '\0' ) break;
			*dest++ = *pxc->param++;
			if ( *pxc->param != ',' ) continue;
			pxc->param++;
			break;
		}
		*dest = '\0';
		if ( buf[0] == '\0' ) break;
		if ( (buf[0] >= '0') && (buf[0] <= '9') ){
			*StayMode = GetIntNumberW(buf);
		}else{
			wcscpy(InvokeName, buf);
		}
	}
	pxc->param += wcslen(pxc->param) + 1; // ":～" をスキップ
	pxc->paramcount--;
}

void SetResult(PPXAPPINFOW* ppxa, PPXMCOMMANDSTRUCT* pxc, CResult^ value, Object^ evresult)
{
	marshal_context ctx;
	const wchar_t* result;

	if (value->UseResult) {
		// null 等は property Object^ result 内で処理する
		result = ctx.marshal_as<const wchar_t*>(value->result);
	}
	else {
		if ((evresult == nullptr) || (evresult->GetType() == Microsoft::ClearScript::Undefined::typeid)) {
			result = L"";
		}
		else if (evresult->GetType() == bool::typeid) {
			result = safe_cast<bool>(evresult) ? L"-1" : L"0";
		}
		else {
			result = ctx.marshal_as<const wchar_t*>(evresult->ToString());
		}
	}

	if (wcslen(result) < 1000) {
		wcscpy(pxc->resultstring, result);
	}
	else {
		WCHAR* longtext;

		longtext = static_cast<WCHAR*>(HeapAlloc(GetProcessHeap(), 0, (wcslen(result) + 1) * sizeof(WCHAR)));
		if (longtext != NULL) {
			wcscpy(longtext, result);
			ppxa->Function(ppxa, PPXCMDID_LONG_RESULT, longtext);
			HeapFree(GetProcessHeap(), 0, longtext);
		}
	}
}

int RunStayFunction(PPXAPPINFOW* ppxa, PPXMCOMMANDSTRUCT* pxc, CResult^ value, int StayMode, CInstance^ instance, OldPPxInfoStruct *OldInfo, WCHAR *InvokeName)
{
	int argc = pxc->paramcount;
	const WCHAR *param = pxc->param;
	Object^ evresult;

	if ( (argc > 0) &&
		 !((InvokeName[0] != '\0' ) &&
			 (StayMode >= ScriptStay_Stay)) ){
		param += wcslen(param) + 1; // 次のパラメータに
		argc--;
	}

	if ( InvokeName[0] == '\0' ) InvokeName = L"ppx_resume";

	array<Object^>^ argv = gcnew array<Object^>(argc);
	for ( int i = 0; i < argc; i++ ){
		argv[i] = marshal_as<String^>(param);
		param += wcslen(param) + 1;
	}
	evresult = instance->engine->Invoke(marshal_as<String^>(InvokeName), argv);
	SleepStayInstance(instance, OldInfo);

	if ( pxc->resultstring != NULL ) SetResult(ppxa, pxc, value, evresult);
	return PPXMRESULT_DONE;
}

int RunScript(PPXAPPINFOW* ppxa, PPXMCOMMANDSTRUCT* pxc, int file)
{
	String^ source;

	if (pxc->paramcount < 1) {
		::PopupMessage(ppxa, (pxc->resultstring == NULL) ?
				"*js \"code\"[,param1...] / *script filename[,param1...]" :
				"%*js(\"code\"[,param1...]) / %*script(filename[,param1...])");
		return PPXMRESULT_STOP;
	}

	CResult^ value = gcnew CResult;

	PPXMCOMMANDSTRUCT pxcbuf;
	int StayMode = ScriptStay_Cache;
	WCHAR InvokeName[256];

	InvokeName[0] = '\0';
	if ( (pxc->paramcount > 0) && (pxc->param[0] == ':') ){
		pxcbuf = *pxc;
		pxc = &pxcbuf;
		CheckOption(pxc, &StayMode, InvokeName);
	}
	try {
		CInstance^ instance;
		Object^ evresult;
		OldPPxInfoStruct OldInfo;

		if (gc_StayInstance != nullptr) {
			instance = GetStayInstance(ppxa, pxc, &OldInfo, StayMode, file);
			if (instance != nullptr) {
				// invoke 実行(関数名有り or インスタンス指定無し)
				if ( (InvokeName[0] != '\0') || (StayMode < ScriptStay_Stay) || (pxc->paramcount == 0) || (pxc->param[0] == '\0') ){
					return RunStayFunction(ppxa, pxc, value, StayMode, instance, &OldInfo, InvokeName);
				}
				// 既存 instance を使用
			}
			else {
				instance = gcnew CInstance(ppxa, pxc, value, StayMode);
			}
		}
		else {
			instance = gcnew CInstance(ppxa, pxc, value, StayMode);
		}
		if ( InvokeName[0] != '\0' ){
			::PopupMessage(ppxa, "not found instance");
			return PPXMRESULT_STOP;
		}
		if (pxc->paramcount == 0) {
			::PopupMessage(ppxa, "no source parameter");
			return PPXMRESULT_STOP;
		}

		String^ sourcetext = InitScriptText(instance, source, file);

		DocumentInfo codedocs(
				(pxc->resultstring == NULL) ? "command" : "function" );
		try {
			evresult = instance->engine->Evaluate(codedocs, sourcetext);
		}
		catch (Exception^ e) {
			if (e->Message->Contains("use import")) {
				instance->info->ModuleMode = 1;
				codedocs.Category = JavaScript::ModuleCategory::Standard;
				evresult = instance->engine->Evaluate(codedocs, sourcetext);
			}
			else {
				throw(e);
			}
		}
		if ( pxc->resultstring != NULL ) SetResult(ppxa, pxc, value, evresult);

		if (instance->info->stay.mode >= ScriptStay_Stay) {
			if (instance->info->stay.threadID == 0) { // 未登録なので登録
				instance->source = source;
				ChainStayInstance(instance);
				ppxa->Function(ppxa, PPXCMDID_REQUIRE_CLOSETHREAD, 0);
			}else{
				SleepStayInstance(instance, &OldInfo);
			}
		}else{
			instance->info->stay.entry--;
			if ( instance->info->stay.entry > 0 ){
				SleepStayInstance(instance, &OldInfo);
			}
		}
		return PPXMRESULT_DONE;
	}
	catch (Exception^ e) {
		int funcresult = AfterException(e, ppxa, source);
		if ( (funcresult == PPXMRESULT_DONE) && (pxc->resultstring != NULL) ){
			SetResult(ppxa, pxc, value, nullptr);
		}
		return funcresult;
	}
}
