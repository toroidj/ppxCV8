/*-----------------------------------------------------------------------------
	Paper Plane xUI  Clear Script V8 Module
-----------------------------------------------------------------------------*/
#pragma once
#define TOSTRMACRO(item)	#item

#define SCRIPTMODULEVER		4  // Release number
#define SCRIPTMODULEVERSTR	UNICODESTR(TOSTRMACRO(4))

using namespace System;
using namespace System::Collections;
using namespace System::Runtime::InteropServices;
using namespace Microsoft::ClearScript;
using namespace Microsoft::ClearScript::V8;
using namespace msclr::interop;

#ifdef VC_DLL_EXPORTS
#undef VC_DLL_EXPORTS
#define VC_DLL_EXPORTS extern "C" __declspec(dllexport)
#else
#define VC_DLL_EXPORTS extern "C" __declspec(dllimport)
#endif

VC_DLL_EXPORTS int WINAPI ModuleEntry(PPXAPPINFOW* ppxa, DWORD cmdID, PPXMODULEPARAM pxs);
