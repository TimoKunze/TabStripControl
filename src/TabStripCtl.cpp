// TabStripCtl.cpp: Implementation of DLL exports.

#include "stdafx.h"
#include "res/resource.h"
#ifdef UNICODE
	#include "TabStripCtlU.h"
#else
	#include "TabStripCtlA.h"
#endif


class TabStripCtlModule :
    public CAtlDllModuleT<TabStripCtlModule>
{

public:
  #ifdef UNICODE
		DECLARE_LIBID(LIBID_TabStripCtlLibU)
		DECLARE_REGISTRY_APPID_RESOURCEID(IDR_TABSTRIPCTL, "{4956FC30-4273-44f5-B144-E57156D61F7A}")
	#else
		DECLARE_LIBID(LIBID_TabStripCtlLibA)
		DECLARE_REGISTRY_APPID_RESOURCEID(IDR_TABSTRIPCTL, "{8D644723-BF6D-436a-A58E-9B872F2C850B}")
	#endif
};

TabStripCtlModule _AtlModule;


#ifdef _MANAGED
	#pragma managed(push, off)
#endif

// DLL entry point
extern "C" BOOL WINAPI DllMain(HINSTANCE /*hInstance*/, DWORD dwReason, LPVOID lpReserved)
{
	#ifdef _DEBUG
		// enable CRT memory leak detection & report
		_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF     // enable debug heap allocs & block type IDs (ie _CLIENT_BLOCK)
		    //| _CRTDBG_CHECK_CRT_DF              // check CRT allocations too
		    //| _CRTDBG_DELAY_FREE_MEM_DF       // keep freed blocks in list as _FREE_BLOCK type
		    | _CRTDBG_LEAK_CHECK_DF             // do leak report at exit (_CrtDumpMemoryLeaks)

		    // pick only one of these heap check frequencies
		    //| _CRTDBG_CHECK_ALWAYS_DF         // check heap on every alloc/free
		    //| _CRTDBG_CHECK_EVERY_16_DF
		    //| _CRTDBG_CHECK_EVERY_128_DF
		    //| _CRTDBG_CHECK_EVERY_1024_DF
		    | _CRTDBG_CHECK_DEFAULT_DF          // by default, no heap checks
		    );
		//_CrtSetBreakAlloc(209);               // break debugger on numbered allocation
		// get ID number from leak detector report of previous run
	#endif

	return _AtlModule.DllMain(dwReason, lpReserved); 
}

#ifdef _MANAGED
	#pragma managed(pop)
#endif

STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer(void)
{
	return _AtlModule.DllRegisterServer();
}

STDAPI DllUnregisterServer(void)
{
	return _AtlModule.DllUnregisterServer();
}