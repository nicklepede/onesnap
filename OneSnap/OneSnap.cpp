// OneSnap.cpp : Implementation of DLL Exports.


#include "stdafx.h"
#include "resource.h"
#include "OneSnap.h"


class COneSnapModule : public CAtlDllModuleT< COneSnapModule >
{
public :
	DECLARE_LIBID(LIBID_OneSnapLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_ONESNAP, "{BC626B2A-B723-49D6-9118-E4862BC511B6}")
};

COneSnapModule _AtlModule;

class COneSnapApp : public CWinApp
{
public:

// Overrides
    virtual BOOL InitInstance();
    virtual int ExitInstance();

    DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(COneSnapApp, CWinApp)
END_MESSAGE_MAP()

COneSnapApp theApp;

BOOL COneSnapApp::InitInstance()
{
    return CWinApp::InitInstance();
}

int COneSnapApp::ExitInstance()
{
#ifdef _DEBUG
	_CrtDumpMemoryLeaks();
#endif
    return CWinApp::ExitInstance();
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState());
    return (AfxDllCanUnloadNow()==S_OK && _AtlModule.GetLockCount()==0) ? S_OK : S_FALSE;
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
	return hr;
}

