// Snapper.h : Declaration of the CSnapper

#pragma once
#include "resource.h"       // main symbols

#include "OneSnap.h"

class _com_error;


// CSnapper
class ATL_NO_VTABLE CSnapper :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSnapper, &CLSID_Snapper>,
	public IObjectWithSiteImpl<CSnapper>,
	public IOleCommandTarget,
	public IDispatchImpl<ISnapper, &IID_ISnapper, &LIBID_OneSnapLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CSnapper()
	{
	}

	STDMETHOD(SetSite)(IUnknown *pUnkSite);
	IWebBrowser2* GetWebBrowser(IUnknown* pUnkSite);

	virtual /* [input_sync] */ HRESULT STDMETHODCALLTYPE QueryStatus( 
		/* [unique][in] */ const GUID *pguidCmdGroup,
		/* [in] */ ULONG cCmds,
		/* [out][in][size_is] */ OLECMD prgCmds[  ],
		/* [unique][out][in] */ OLECMDTEXT *pCmdText);

		virtual HRESULT STDMETHODCALLTYPE Exec( 
		/* [unique][in] */ const GUID *pguidCmdGroup,
		/* [in] */ DWORD nCmdID,
		/* [in] */ DWORD nCmdexecopt,
		/* [unique][in] */ VARIANT *pvaIn,
		/* [unique][out][in] */ VARIANT *pvaOut);


	HRESULT CSnapper::Exec(void);



		DECLARE_REGISTRY_RESOURCEID(IDR_SNAPPER)

	BEGIN_COM_MAP(CSnapper)
		COM_INTERFACE_ENTRY(ISnapper)
		COM_INTERFACE_ENTRY(IObjectWithSite)
		COM_INTERFACE_ENTRY(IOleCommandTarget)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

private:

	BOOL CapturePage(BYTE** ppbImage, long* pcbImage, long* plHeight, long* plWidth);
	BOOL CapturePage(LPCTSTR pszFilename, long* plHeight, long* plWidth);
	HRESULT GetBrowserWindow(HWND* hWnd);
	BOOL CreateBitmapImage(HWND hWnd, HBITMAP hBMP, HDC hDC, BYTE** ppbImage, long* pcbImage, long* plHeight, long* plWidth);
	PBITMAPINFO CreateBitmapInfoStruct(HWND hwnd, HBITMAP hBmp);
	BOOL CreateBMPFile(HWND hwnd, LPTSTR pszFile, HBITMAP hBMP, HDC hDC);
	BOOL WriteDataFile(LPTSTR pszFile, BYTE* pbData, DWORD cbSize, HWND hwnd);
	void OneNoteErrorMsg(CAtlString& strErr, _com_error& e);
	CAtlString	GetPageTitle();
	CAtlString GetPageHtml();
	CAtlString	GetPageUrl();
	void GetDpi(int* pdpiVert, int* pdpiHoriz);

	CComPtr<IWebBrowser2> m_spBrowser;
};

OBJECT_ENTRY_AUTO(__uuidof(Snapper), CSnapper)
