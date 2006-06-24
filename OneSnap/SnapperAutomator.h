// SnapperAutomator.h : Declaration of the CSnapperAutomator

#pragma once
#include "resource.h"       // main symbols

#include "OneSnap.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CSnapperAutomator

class ATL_NO_VTABLE CSnapperAutomator :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSnapperAutomator, &CLSID_SnapperAutomator>,
	public IObjectWithSiteImpl<CSnapperAutomator>,
	public IDispatchImpl<ISnapperAutomator, &IID_ISnapperAutomator, &LIBID_OneSnapLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CSnapperAutomator()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_SNAPPERAUTOMATOR)


BEGIN_COM_MAP(CSnapperAutomator)
	COM_INTERFACE_ENTRY(ISnapperAutomator)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IObjectWithSite)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	STDMETHOD(SnapIt)(IUnknown* pUnknown);
};

OBJECT_ENTRY_AUTO(__uuidof(SnapperAutomator), CSnapperAutomator)
