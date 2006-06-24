// OneSnap.cpp : Implementation of CSnapper
#include "stdafx.h"

#include "OneSnap.h"
#include "Snapper.h"
#include "SnapperDialog.h"
#include "SnapperConfig.h"
#import "progid:OneNote.CSimpleImporter.1"

using namespace OneNote;
/*********************************

Templated items are:

	- PATH
	- SECTION
	- PAGE_TITLE
	- PAGE_GUID
	- PAGE_PATH
	- IMAGE_GUID
	- IMAGE_FILEPATH
	- IMAGE_POSX
	- IMAGE_POSY 
	- IMAGE_HEIGHT
	- IMAGE_WIDTH
	- IMAGE_BACKGROUND
	- HEADER_XML
	- HEADER_HTML
	- HEADER_POSX
	- HEADER_POSY
	- HEADER_WIDTH
	- LINK
	- COMMENT
**********************************/

#define IMPORT_XML	"\
<?xml version=\"1.0\" encoding=\"utf-16\"?>\
<Import xmlns=\"http://schemas.microsoft.com/office/onenote/2004/import\">\
<EnsurePage path=\"<PAGE_PATH>\" guid=\"<PAGE_GUID>\" title=\"<TITLE>\" />\
<PlaceObjects pagePath=\"<PAGE_PATH>\" pageGuid=\"<PAGE_GUID>\">\
<HEADER_XML>\
<Object guid=\"<IMAGE_GUID>\">\
<Position x=\"<IMAGE_POSX>\" y=\"<IMAGE_POSY>\" />\
<Image width=\"<IMAGE_WIDTH>\" height=\"<IMAGE_HEIGHT>\" backgroundImage=\"<IMAGE_BACKGROUND>\">\
<File path=\"<IMAGE_FILEPATH>\" />\
</Image>\
</Object>\
</PlaceObjects>\
</Import>\
"

#define PAGEHTML_XML	"\
<Object guid=\"<PAGEHTML_GUID>\">\
<Position x=\"<PAGEHTML_POSX>\" y=\"<PAGEHTML_POSY>\" />\
<Outline width=\"<PAGEHTML_WIDTH>\">\
<Html><Data><![CDATA[<PAGEHTML_HTML>]]></Data></Html>\
</Outline>\
</Object>\
"

#define HEADER_XML	"\
<Object guid=\"<HEADER_GUID>\">\
<Position x=\"<HEADER_POSX>\" y=\"<HEADER_POSY>\" />\
<Outline width=\"<HEADER_WIDTH>\">\
<Html><Data><![CDATA[<HEADER_HTML>]]></Data></Html>\
</Outline>\
</Object>\
"

#define HEADER_HTML	"\
<html>\
<body>\
<FONT SIZE=4><a href=\"<LINK>\"><span style=\"font-weight: bold;\"><TITLE></span></a><br><br></FONT>\
<COMMENT><br><br><span style=\"font-weight: bold;\"><span style=\"font-weight: bold;\"></span><br><br></span>\
<hr style=\"width: 100%; height: 2px;\"><span style=\"font-weight: bold;\"></span>\
</body>\
</html>\
"

#define PAGEHTML_STUB	"\
<html>\
<body>\
<FONT SIZE=4><a href=\"http:\\www.hotmail.com\"><span style=\"font-weight: bold;\">hello</span></a><br><br></FONT>\
comment<br><br><span style=\"font-weight: bold;\"><span style=\"font-weight: bold;\"></span><br><br></span>\
<hr style=\"width: 100%; height: 2px;\"><span style=\"font-weight: bold;\"></span>\
</body>\
</html>\
"


#define COM_TRY(fcn) { HRESULT _hr; if FAILED(_hr = (fcn)) _com_issue_error(_hr); }
/////////////////////////////////////////////////////////////////////////////
// CSnapper



HRESULT CSnapper::Exec(void)
{
	AFX_MANAGE_STATE( AfxGetStaticModuleState( ));
	try
	{
		long lHeight, lWidth;
		BOOL bSucceeded = FALSE;

		CString strUrl = GetPageUrl();
		CSnapperConfig	cfg;
/*		CWnd		cwndBrowser;
		cwndBrowser.Attach(GetBrowserWindow(&hWnd))*/
		CSnapperDialog	dlgConfig;
		UserSettings	sSettings;

		sSettings.strTitle = GetPageTitle();
		if (IDOK == dlgConfig.DoModalEx(&sSettings, &cfg))
		{
			// sleep for a bit just to make sure the modal dialog is gone.  (You'd think it would already be gone, but it's not)
			Sleep(1);
			// setup a temporary scratch file for capture to use...
			CAtlTemporaryFile	fTemp;
			COM_TRY(fTemp.Create(NULL, GENERIC_READ | GENERIC_WRITE));
			COM_TRY(fTemp.HandsOff());

			LPCTSTR pszFilename = fTemp.TempFileName();

			if (!CapturePage(pszFilename, &lHeight, &lWidth)) throw "capture failed";

			CAtlString  strImport(TEXT(IMPORT_XML));

			GUID	idPage;
			if (FAILED(CoCreateGuid(&idPage))) throw "error creating guid";
			_bstr_t  bstrGuid(TEXT("{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"));

			if (0 == StringFromGUID2(idPage, bstrGuid, bstrGuid.length()+1)) throw "error formatting guid";
			strImport.Replace(TEXT("<PAGE_GUID>"), bstrGuid);

			GUID	idObjImage;
			if (FAILED(CoCreateGuid(&idObjImage))) throw "error creating guid";
			_bstr_t  bstrImageGuid(TEXT("{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"));
			if (0 == StringFromGUID2(idObjImage, bstrImageGuid, bstrImageGuid.length()+1)) throw "error formatting guid";
			strImport.Replace(TEXT("<IMAGE_GUID>"), bstrImageGuid);
			CAtlString  strHeight;
			// replace all the template words w/ the real thing.
			strImport.Replace(TEXT("<TITLE>"), sSettings.strTitle);
			strImport.Replace(TEXT("<PAGE_PATH>"), sSettings.strSectionFilepath);
			strImport.Replace(TEXT("<IMAGE_POSX>"), TEXT("36"));
			strImport.Replace(TEXT("<IMAGE_POSY>"), TEXT("72"));
			CAtlString  strWidth;

			int dpiVert = 96;
			int dpiHoriz = 96;
			GetDpi(&dpiVert, &dpiHoriz);
			// determine the size of the image in inches
			float fHeightInches = (static_cast<float>(lHeight)) / ((float) dpiVert);
			float fWidthInches  = (static_cast<float>(lWidth)) / ((float) dpiHoriz);
			// if the user wants us to limit the page width then scale, if needed.
			if (cfg.GetLimitPageWidth())
			{
				float fMaxWidth = cfg.GetMaxPageWidth();
				if (fWidthInches > fMaxWidth)
				{
					// image is too wide, so scale it down to the max width
					float fScale = fMaxWidth / fWidthInches;
					fWidthInches *= fScale;
					fHeightInches *= fScale;
				}
			}

			// The OneNote import API's coordinates are based on 72 dots per inch, so scale accordingly
			long lHeightResized = static_cast<long> (fHeightInches * 72.0f);
			strHeight.Format(TEXT("%d"),lHeightResized);
			strImport.Replace(TEXT("<IMAGE_HEIGHT>"), strHeight);

			long lWidthResized = static_cast<long> (fWidthInches * 72.0f);
			strWidth.Format(TEXT("%d"), lWidthResized); 
			strImport.Replace(TEXT("<IMAGE_WIDTH>"), strWidth);
			strImport.Replace(TEXT("<IMAGE_FILEPATH>"), pszFilename); 
			LPTSTR pszImportAsBackground = _T("false");
			if (cfg.GetImportAsBackground())
				pszImportAsBackground = _T("true");
			strImport.Replace(_T("<IMAGE_BACKGROUND>"), pszImportAsBackground);


			// add header
			CAtlString strHeader(TEXT(HEADER_XML));

			GUID	idObjHeader;
			_bstr_t  bstrHeaderGuid(TEXT("{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"));
			if (FAILED(CoCreateGuid(&idObjHeader))) throw "error creating header guid";
			if (0 == StringFromGUID2(idObjHeader, bstrHeaderGuid, bstrHeaderGuid.length()+1)) throw "error formatting guid";
			strHeader.Replace(TEXT("<HEADER_GUID>"), bstrHeaderGuid);

			strHeader.Replace(TEXT("<HEADER_POSX>"), TEXT("36"));
			strHeader.Replace(TEXT("<HEADER_POSY>"), TEXT("10"));
			strHeader.Replace(TEXT("<HEADER_WIDTH>"), TEXT("500"));

			CAtlString strHtml(TEXT(HEADER_HTML));
			strHtml.Replace(TEXT("<LINK>"), strUrl);
			strHtml.Replace(TEXT("<TITLE>"), sSettings.strTitle);
			strHtml.Replace(TEXT("<COMMENT>"), sSettings.strComment );
			strHeader.Replace(TEXT("<HEADER_HTML>"), strHtml);
			strImport.Replace(TEXT("<HEADER_XML>"), strHeader);
#if 0
			// add page html
			// removing for now because OneNote seems to be fairly picky on the HTML it accepts.  
			if (TRUE)
			{
				// create an object guid.
				GUID	idPageHtml;
				_bstr_t  bstrPageHtmlGuid(TEXT("{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"));
				if (FAILED(CoCreateGuid(&idPageHtml))) throw "error creating html page guid";
				if (0 == StringFromGUID2(idPageHtml, bstrPageHtmlGuid, bstrPageHtmlGuid.length()+1)) throw "error formatting guid";

				CAtlString strPageHtml(_T(PAGEHTML_XML));
				strPageHtml.Replace(_T("<PAGEHTML_GUID>"), bstrPageHtmlGuid);
				strPageHtml.Replace(_T("<PAGEHTML_POSX>"), _T("72"));
				long lPageHtmlPosY = lHeightResized + 36;
				CString strPageHtmlPosY;
				strPageHtmlPosY.Format(_T("%d"), lPageHtmlPosY);
				strPageHtml.Replace(_T("<PAGEHTML_POSY>"), strPageHtmlPosY);
				strPageHtml.Replace(_T("<PAGEHTML_WIDTH>"), _T("500"));
				strPageHtml.Replace(_T("<PAGEHTML_HTML>"), /*_T(PAGEHTML_STUB)*/ GetPageHtml());
				strImport.Replace(_T("<PAGEHTML_XML>"), strPageHtml);
			}
			else
			{
				strImport.Replace(_T("<PAGEHTML_XML>"), _T(""));
			}
#endif
			OutputDebugString(strImport);

			ISimpleImporterPtr spOneNote(__uuidof(CSimpleImporter));
			_bstr_t bstrImport(strImport);
			spOneNote->Import(bstrImport);
			fTemp.HandsOn();
			
			// navigate to the page if the option is set
			if (cfg.GetNavigateToPage())
			{
				_bstr_t bstrPath(sSettings.strSectionFilepath);
				spOneNote->NavigateToPage(bstrPath , bstrGuid);
			}
			bSucceeded = TRUE;
		}
	}
	catch (_com_error e)
	{
		CAtlString strErr;
		OneNoteErrorMsg(strErr, e);

		MessageBox(NULL, strErr, TEXT("OneSnap Error!"), MB_OK | MB_ICONERROR);	
	}
	catch (TCHAR* pszMsg)
	{
		MessageBox(NULL, pszMsg, _T("OneSnap Error!"), MB_OK | MB_ICONERROR);
	}
	catch (CString strMsg)
	{
		MessageBox(NULL, strMsg, _T("OneSnap Error!") , MB_OK | MB_ICONERROR);		
	}
	catch(...)
	{
		MessageBox(NULL, TEXT("Error during OneNote import!"), TEXT("OneSnap Error!"), MB_OK | MB_ICONERROR);
	}

	return S_OK;

}

void CSnapper::GetDpi(int* pdpiVert, int* pdpiHoriz)
{
	HDC screen = GetDC(0);

   *pdpiHoriz = GetDeviceCaps(screen, LOGPIXELSX);
   *pdpiVert = GetDeviceCaps(screen, LOGPIXELSY);
   ReleaseDC(0, screen);
}

STDMETHODIMP CSnapper::Exec ( 
/* [unique][in] */ const GUID __RPC_FAR *pguidCmdGroup,
/* [in] */ DWORD nCmdID,
/* [in] */ DWORD nCmdexecopt,
/* [unique][in] */ VARIANT __RPC_FAR *pvaIn,
/* [unique][out][in] */ VARIANT __RPC_FAR *pvaOut)
{ 

	return Exec();
}

CAtlString CSnapper::GetPageTitle()
{

	CComQIPtr<IHTMLDocument2> spDoc;
	CComPtr<IDispatch> spDocDispatch;

	COM_TRY(m_spBrowser->get_Document(&spDocDispatch))

	spDoc = spDocDispatch;
	if (spDoc == NULL) throw _T("problem getting IHTMLDocument2 interface from browser object");

	VARIANT varTitle;
	V_VT(&varTitle) = VT_BSTR;
	COM_TRY(spDoc->get_title(&varTitle.bstrVal));

	CAtlString strTitle(varTitle);
	return strTitle;
}
CAtlString CSnapper::GetPageHtml()
{

	CComQIPtr<IHTMLDocument3> spDoc;
	CComPtr<IDispatch> spDocDispatch;
	CComPtr<IHTMLElement>	spDocElement;
	COM_TRY(m_spBrowser->get_Document(&spDocDispatch))

	spDoc = spDocDispatch;
	if (spDoc == NULL) throw _T("problem getting IHTMLDocument3 interface from browser object");
	
	COM_TRY(spDoc->get_documentElement(&spDocElement));

	VARIANT varPage;
	V_VT(&varPage) = VT_BSTR;
	COM_TRY(spDocElement->get_outerHTML(&varPage.bstrVal));

	CAtlString strPage(varPage);
	return strPage;
}
CAtlString CSnapper::GetPageUrl()
{
	VARIANT varUrl;
	V_VT(&varUrl) = VT_BSTR;
	COM_TRY(m_spBrowser->get_LocationURL(&varUrl.bstrVal));

	CAtlString strUrl(varUrl);
	return strUrl;
}
BOOL CSnapper::CapturePage(LPCTSTR pszFilename, long* plHeight, long* plWidth)
{

	BYTE* pbImage = NULL;
	long  cbImage;

	BOOL bSucceeded = FALSE;

	if (CapturePage(&pbImage, &cbImage, plHeight, plWidth))
	{
		// write the file to disk
		HWND hWnd = NULL;
		if (SUCCEEDED(GetBrowserWindow(&hWnd)))
		{
			if (WriteDataFile((LPTSTR) pszFilename, pbImage, cbImage, hWnd))
			{
				bSucceeded = TRUE;
			}
		}
	}

	if (pbImage != NULL)
		GlobalFree(pbImage);

	return bSucceeded;

}
BOOL CSnapper::CapturePage(BYTE** ppbImage, long* pcbImage, long* plHeight, long* plWidth)
{
	BOOL bSucceeded = FALSE;
	HDC hdcMain = NULL;
	HDC hdcCom = NULL;
	HWND hWnd = NULL;
	long lScrollHeight, lScrollWidth;
	long lScrollTopStart, lScrollTopCur;
	long lClientHeight, lClientWidth;
	long lHtmlHeight, lHtmlWidth;
	long lHtmlScrollHeight, lHtmlScrollWidth;
	long lPage;
	long i;
	BYTE* pbImage = NULL;
	long cbImage = 0;

	RECT rcClient;	  
	HWND hWndDocObj = NULL;
	HWND hWndExpServer = NULL;
	HBITMAP hbmSnap = NULL;
	HGDIOBJ hbmPrevious = NULL;
	CComQIPtr<IHTMLDocument2> spDoc;
	CComPtr<IHTMLElement> spBody;
	CComQIPtr<IHTMLElement2> spBody2;
	CComQIPtr<IViewObject> spView;
	CComQIPtr<IHTMLElementRender> spRender;
	CComPtr<IDispatch> spDocDispatch;
	CComQIPtr<IHTMLElement> spHtml;
	CComQIPtr<IHTMLElement2> spHtml2;
	CComQIPtr<IHTMLDocument3> spDoc3;

	if (m_spBrowser == NULL) 
		return NULL;
	//m_spBrowser->get_LocationURL(pbstrUrl);

	m_spBrowser->get_Document(&spDocDispatch);
	if (spDocDispatch == NULL)  goto cleanup;

	spDoc = spDocDispatch;
	if (spDoc == NULL) goto cleanup;

//	if (FAILED(spDoc->get_title(pbstrTitle))) goto cleanup;

	spDoc3 = spDoc;
	if (spDoc3 == NULL) goto cleanup;
	spDoc3->get_documentElement(&spHtml);
	if (spHtml == NULL) goto cleanup;
	spHtml2 = spHtml; 
	if (spHtml2 == NULL) goto cleanup;
	if (FAILED(spHtml2->get_clientHeight(&lHtmlHeight))) goto cleanup;
	if (FAILED(spHtml2->get_clientWidth(&lHtmlWidth))) goto cleanup;
	if (FAILED(spHtml2->get_scrollHeight(&lHtmlScrollHeight))) goto cleanup;
	if (FAILED(spHtml2->get_scrollWidth(&lHtmlScrollWidth))) goto cleanup;

	spDoc->get_body(&spBody);
	if (spBody == NULL) goto cleanup;

	spBody2 = spBody;
	if (spBody2 == NULL) goto cleanup;

	spRender = spBody;
	if (spBody2 == NULL) goto cleanup;

	if (FAILED(spBody2->get_scrollWidth(&lScrollWidth))) goto cleanup;
	if (FAILED(spBody2->get_scrollHeight(&lScrollHeight))) goto cleanup;
	if (FAILED(spBody2->get_scrollTop(&lScrollTopStart))) goto cleanup;

	if (FAILED(spBody2->get_clientHeight(&lClientHeight))) goto cleanup;
	if (FAILED(spBody2->get_clientWidth(&lClientWidth))) goto cleanup;

	if (FAILED(GetBrowserWindow(&hWnd))) goto cleanup;

	// repaint the window, just in case some old dialog box remnants are still around...
	UpdateWindow(hWnd); 


	hWndDocObj = FindWindowEx(hWnd, NULL, TEXT("Shell DocObject View"), NULL);
	if (hWndDocObj == NULL) goto cleanup;

	hWndExpServer = FindWindowEx(hWndDocObj, NULL, TEXT("Internet Explorer_Server"), NULL);
	if (hWndExpServer == NULL) goto cleanup;

	GetClientRect(hWndExpServer, &rcClient);

	if (lClientHeight > (rcClient.bottom - rcClient.top))
	{
		// sometime the root HTML element in the DOM model contains the scrollable page data.
		// (e.g., http://www.bloomberg.com/apps/news?pid=10000102&sid=atM1zE4X3X2c&refer=uk)
		// So if we *think* this is the case then use the HTML element instead of the body.
		
		lClientHeight = lHtmlHeight;
		lClientWidth = lHtmlWidth;
		lScrollHeight = lHtmlScrollHeight;
		lScrollWidth = lHtmlScrollWidth;

		spBody2 = spHtml2;

	}


	hdcMain = GetDC(hWndExpServer); 
	if (hdcMain == NULL) goto cleanup;

	hdcCom = CreateCompatibleDC(hdcMain); 
	if (hdcCom == NULL) goto cleanup;

	lPage = 0;
	//Get Screen Height (for bottom up screen drawing)
	while ((lPage * lClientHeight) < lScrollHeight)
    {
        if (FAILED(spBody2->put_scrollTop( (lClientHeight - 5)*lPage)))
			goto cleanup;
		lPage++;
    }
    //Rollback the page count by one
    lPage--;

#define TRIM 3
	// create a large (24-bit) bitmap to blit the entire web page into.
	//hbmSnap = CreateCompatibleBitmap(hdcMain, lClientWidth - TRIM, lScrollHeight - (TRIM * lPage));
	BITMAPINFOHEADER bmi;
	ZeroMemory(&bmi, sizeof(BITMAPINFOHEADER));
	bmi.biSize = sizeof(BITMAPINFOHEADER);
	bmi.biBitCount = 24;
	bmi.biHeight = lScrollHeight - (TRIM * (lPage + 1));
	bmi.biWidth = lClientWidth - TRIM;
	bmi.biPlanes = 1;
	bmi.biCompression  = BI_RGB;

	void* pvBits = NULL;
	hbmSnap = CreateDIBSection(hdcMain, (BITMAPINFO*) &bmi, DIB_RGB_COLORS, &pvBits, NULL, 0);
	if (hbmSnap == NULL)
	{
		DWORD dwErr = GetLastError();
		goto cleanup;
	}

	/* now make the bitmap by blting the image to it */ 
	hbmPrevious = SelectObject(hdcCom, hbmSnap);
	if (hbmPrevious = NULL) goto cleanup;

	for (i = lPage; i >= 0; --i)
    {
        //Shoot visible window
        if (FAILED(spBody2->put_scrollTop( (lClientHeight - 5)* i))) goto cleanup;
		if (FAILED(spBody2->get_scrollTop( &lScrollTopCur ))) goto cleanup;

		//spRender->SetDocumentPrinter(_bstr_t("foo"), hdcCom);
		//spRender->DrawToDC(hdcCom);
		if (!BitBlt(hdcCom, 0, lScrollTopCur, lClientWidth - TRIM, lClientHeight - TRIM, hdcMain, TRIM,TRIM, SRCCOPY)) goto cleanup;
    }

	// reset the scroll window.
	if (FAILED(spBody2->put_scrollTop(lScrollTopStart))) goto cleanup;

	//if (!CreateBMPFile(hWnd, "C:\\foo3.bmp", hbmSnap, hdcCom) ) goto cleanup;
	if (!CreateBitmapImage(hWnd, hbmSnap, hdcCom, &pbImage, &cbImage, plHeight, plWidth) ) goto cleanup;

	bSucceeded = TRUE;

cleanup:

	if (hbmPrevious != NULL && hdcCom != NULL)
		SelectObject(hdcCom, hbmPrevious);

	if (hbmSnap != NULL)
		DeleteObject(hbmSnap);

	if (hdcCom != NULL)
		DeleteDC(hdcCom);

	if (hdcMain != NULL)
		ReleaseDC(hWnd, hdcMain);

	if (!bSucceeded)
	{
//		::SysFreeString(*pbstrTitle);
//		::SysFreeString(*pbstrUrl);
		if (hWnd == NULL)
			hWnd = GetForegroundWindow();
		MessageBox(hWnd, _T("Error - couldn't capture page"), _T("OneSnap Error!"), MB_ICONERROR);
	}
	*ppbImage = pbImage;
	*pcbImage = cbImage;
	return bSucceeded;
}



STDMETHODIMP CSnapper::SetSite(IUnknown *pUnkSite)
{
	if (!pUnkSite)
	{
		ATLTRACE(_T("SetSite(): pUnkSite is NULL\n"));
		m_spBrowser = NULL;
	}
	else
	{
		// Get the web browser's external dispatch interface.
		m_spBrowser = GetWebBrowser(pUnkSite);
	}
	return S_OK;
}

HRESULT CSnapper::GetBrowserWindow(HWND* phWnd)
{
	if (m_spBrowser == NULL)
		return E_UNEXPECTED;
	HRESULT hr;
	if (FAILED(hr = m_spBrowser->get_HWND((long*) phWnd))) 
	{
		// apps other than IE (e.g., FeedDemon) use an embedded IE object that 
		// doesn't support IWebBrowser2::get_HWND, so we must get the window
		// handle some other way...
		CComQIPtr<IOleWindow> spWindow = m_spBrowser;
		if (spWindow != NULL)
			hr = spWindow->GetWindow(phWnd);
		else
			hr = E_POINTER;
	}
	
	return hr;
}


IWebBrowser2* CSnapper::GetWebBrowser(IUnknown* pUnkSite)
{

	if (pUnkSite == NULL)
		return NULL;

	IWebBrowser2* piBrowser = NULL;
	::IServiceProvider *piServiceProvider		= NULL;
	::IServiceProvider *piTopServiceProvider  = NULL;

	HRESULT hr = pUnkSite->QueryInterface(IID_IServiceProvider, reinterpret_cast<void **>(&piServiceProvider));
	if (FAILED(hr))
		goto cleanup;

	hr = piServiceProvider->QueryService(SID_STopLevelBrowser, IID_IServiceProvider, reinterpret_cast<void **>(&piTopServiceProvider));
	if (FAILED(hr))
		goto cleanup;
	
	hr = piTopServiceProvider->QueryService(SID_SWebBrowserApp, IID_IWebBrowser2, reinterpret_cast<void **>(&piBrowser));
	if (FAILED(hr))
		goto cleanup;

cleanup:
	// Free resources.
	if (piServiceProvider != NULL)
		piServiceProvider->Release();

	if (piTopServiceProvider != NULL)
		piTopServiceProvider->Release();

	return piBrowser;

}


PBITMAPINFO CSnapper::CreateBitmapInfoStruct(HWND hwnd, HBITMAP hBmp)
{ 
    BITMAP bmp; 
    PBITMAPINFO pbmi; 
    WORD    cClrBits; 

    // Retrieve the bitmap color format, width, and height. 
    if (!GetObject(hBmp, sizeof(BITMAP), (LPSTR)&bmp)) 
		return NULL;
//        errhandler("GetObject", hwnd); 

    // Convert the color format to a count of bits. 
    cClrBits = (WORD)(bmp.bmPlanes * bmp.bmBitsPixel); 
    if (cClrBits == 1) 
        cClrBits = 1; 
    else if (cClrBits <= 4) 
        cClrBits = 4; 
    else if (cClrBits <= 8) 
        cClrBits = 8; 
    else if (cClrBits <= 16) 
        cClrBits = 16; 
    else if (cClrBits <= 24) 
        cClrBits = 24; 
    else cClrBits = 32; 

    // Allocate memory for the BITMAPINFO structure. (This structure 
    // contains a BITMAPINFOHEADER structure and an array of RGBQUAD 
    // data structures.) 

     if (cClrBits != 24) 
         pbmi = (PBITMAPINFO) GlobalAlloc(LPTR, 
                    sizeof(BITMAPINFOHEADER) + 
                    sizeof(RGBQUAD) * (1<< cClrBits)); 

     // There is no RGBQUAD array for the 24-bit-per-pixel format. 

     else 
         pbmi = (PBITMAPINFO) GlobalAlloc(LPTR, 
                    sizeof(BITMAPINFOHEADER)); 

    // Initialize the fields in the BITMAPINFO structure. 

    pbmi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER); 
    pbmi->bmiHeader.biWidth = bmp.bmWidth; 
    pbmi->bmiHeader.biHeight = bmp.bmHeight; 
    pbmi->bmiHeader.biPlanes = bmp.bmPlanes; 
    pbmi->bmiHeader.biBitCount = bmp.bmBitsPixel; 
    if (cClrBits < 24) 
        pbmi->bmiHeader.biClrUsed = (1<<cClrBits); 

    // If the bitmap is not compressed, set the BI_RGB flag. 
    pbmi->bmiHeader.biCompression = BI_RGB; 

    // Compute the number of bytes in the array of color 
    // indices and store the result in biSizeImage. 
    // For Windows NT, the width must be DWORD aligned unless 
    // the bitmap is RLE compressed. This example shows this. 
    // For Windows 95/98/Me, the width must be WORD aligned unless the 
    // bitmap is RLE compressed.
    pbmi->bmiHeader.biSizeImage = ((pbmi->bmiHeader.biWidth * cClrBits +31) & ~31) /8
                                  * pbmi->bmiHeader.biHeight; 
    // Set biClrImportant to 0, indicating that all of the 
    // device colors are important. 
     pbmi->bmiHeader.biClrImportant = 0; 
     return pbmi; 
 } 

#define errhandler(str,hwnd)   MessageBox(hwnd, TEXT(str), TEXT("ERROR!"), MB_OK)




BOOL CSnapper::CreateBitmapImage(HWND hwnd, HBITMAP hBMP, HDC hDC, BYTE** ppbImage, long* pcbImage, long* plHeight, long* plWidth) 
 { 
    BITMAPFILEHEADER hdr;       // bitmap file-header 
    PBITMAPINFOHEADER pbih;     // bitmap info-header 
    LPBYTE lpBits = NULL;              // memory pointer 
    DWORD dwTotal;              // total count of bytes 
    DWORD cb;                   // incremental count of bytes 
    BYTE *hp;                   // byte pointer 
	PBITMAPINFO pbi = NULL;
	BYTE* pbImage = NULL;
	BYTE* pbWrite = NULL;
	long cbSizeHeader;
	long cbSize = 0;

	BOOL bSuccess = FALSE;
	pbi = CreateBitmapInfoStruct(hwnd, hBMP);
	if (pbi == NULL)
	{
		errhandler("CreateBitmapInfoHeader", hwnd);
		goto cleanup;
	}
    pbih = (PBITMAPINFOHEADER) pbi; 
    lpBits = (LPBYTE) GlobalAlloc(GMEM_FIXED, pbih->biSizeImage);
    if (!lpBits) 
	{
         errhandler("GlobalAlloc lpBits", hwnd); 
		 goto cleanup;
	}
    // Retrieve the color table (RGBQUAD array) and the bits 
    // (array of palette indices) from the DIB. 
    if (!GetDIBits(hDC, hBMP, 0, (WORD) pbih->biHeight, lpBits, pbi, 
        DIB_RGB_COLORS)) 
    {
        errhandler("GetDIBits", hwnd); 
		goto cleanup;
    }

    hdr.bfType = 0x4d42;        // 0x42 = "B" 0x4d = "M" 
    // Compute the size of the entire file. 
    hdr.bfSize = (DWORD) (sizeof(BITMAPFILEHEADER) + 
                 pbih->biSize + pbih->biClrUsed 
                 * sizeof(RGBQUAD) + pbih->biSizeImage); 
    hdr.bfReserved1 = 0; 
    hdr.bfReserved2 = 0;

	cbSize = hdr.bfSize;
    pbImage = (BYTE*) GlobalAlloc(GMEM_FIXED, cbSize);
    if (!pbImage) 
	{
         errhandler("GlobalAlloc pbImage", hwnd); 
		 goto cleanup;
	}
	pbWrite = pbImage;


    // Compute the offset to the array of color indices. 
    hdr.bfOffBits = (DWORD) sizeof(BITMAPFILEHEADER) + 
                    pbih->biSize + pbih->biClrUsed 
                    * sizeof (RGBQUAD); 

    // Copy the BITMAPFILEHEADER into the .BMP file. 
	memcpy(pbWrite, &hdr, sizeof(BITMAPFILEHEADER));
	pbWrite += sizeof(BITMAPFILEHEADER);

	cbSizeHeader = sizeof(BITMAPINFOHEADER) + pbih->biClrUsed * sizeof(RGBQUAD);
	memcpy(pbWrite, pbih, cbSizeHeader);
	pbWrite += cbSizeHeader;

    // Copy the array of color indices into the .BMP file. 
    dwTotal = cb = pbih->biSizeImage; 
    hp = lpBits; 
	memcpy(pbWrite, hp, cb);
	pbWrite += cb;

	ATLASSERT(pbWrite == (pbImage + cbSize));

	*plHeight = pbih->biHeight;
	*plWidth = pbih->biWidth;

	bSuccess = TRUE;
cleanup:
    // Free memory.
    if (lpBits != NULL)
		GlobalFree((HGLOBAL)lpBits);

	if (pbi != NULL)
		GlobalFree(pbi);

	if (!bSuccess)
	{
		if (pbImage != NULL)
		{
			GlobalFree(pbImage);
			pbImage = NULL;
			cbSize = 0;
		}
	}

	*pcbImage = cbSize;
	*ppbImage = pbImage;
	return bSuccess;
}




BOOL CSnapper::WriteDataFile(LPTSTR pszFile, BYTE* pbData, DWORD cbSize, HWND hwnd)
{
	BOOL bSucceeded = FALSE;
    HANDLE hf = NULL;                 // file handle 
	DWORD dwTmp;

    // Create the .BMP file. 
    hf = CreateFile(pszFile, 
                   GENERIC_READ | GENERIC_WRITE, 
                   (DWORD) 0, 
                    NULL, 
                   CREATE_ALWAYS, 
                   FILE_ATTRIBUTE_NORMAL, 
                   (HANDLE) NULL); 
    if (hf == INVALID_HANDLE_VALUE) 
	{
        errhandler("CreateFile", hwnd); 
		goto cleanup;
	}

    if (!WriteFile(hf, (LPSTR) pbData, cbSize, (LPDWORD) &dwTmp,NULL)) 
	{
           errhandler("WriteFile", hwnd); 
		   goto cleanup;
	}
	bSucceeded = TRUE;
cleanup:
	 if (hf != NULL)
		 CloseHandle(hf);
	 return bSucceeded;
}





BOOL CSnapper::CreateBMPFile(HWND hwnd, LPTSTR pszFile,  
                  HBITMAP hBMP, HDC hDC) 
 { 
    HANDLE hf;                 // file handle 
    BITMAPFILEHEADER hdr;       // bitmap file-header 
    PBITMAPINFOHEADER pbih;     // bitmap info-header 
    LPBYTE lpBits = NULL;              // memory pointer 
    DWORD dwTotal;              // total count of bytes 
    DWORD cb;                   // incremental count of bytes 
    BYTE *hp;                   // byte pointer 
    DWORD dwTmp; 
	PBITMAPINFO pbi = NULL;

	BOOL bSuccess = FALSE;
	pbi = CreateBitmapInfoStruct(hwnd, hBMP);
	if (pbi == NULL)
	{
		errhandler("CreateBitmapInfoHeader", hwnd);
		goto cleanup;
	}
    pbih = (PBITMAPINFOHEADER) pbi; 
    lpBits = (LPBYTE) GlobalAlloc(GMEM_FIXED, pbih->biSizeImage);

    if (!lpBits) 
	{
         errhandler("GlobalAlloc", hwnd); 
		 goto cleanup;
	}
    // Retrieve the color table (RGBQUAD array) and the bits 
    // (array of palette indices) from the DIB. 
    if (!GetDIBits(hDC, hBMP, 0, (WORD) pbih->biHeight, lpBits, pbi, 
        DIB_RGB_COLORS)) 
    {
        errhandler("GetDIBits", hwnd); 
		goto cleanup;
    }

    // Create the .BMP file. 
    hf = CreateFile(pszFile, 
                   GENERIC_READ | GENERIC_WRITE, 
                   (DWORD) 0, 
                    NULL, 
                   CREATE_ALWAYS, 
                   FILE_ATTRIBUTE_NORMAL, 
                   (HANDLE) NULL); 
    if (hf == INVALID_HANDLE_VALUE) 
	{
        errhandler("CreateFile", hwnd); 
		goto cleanup;
	}
    hdr.bfType = 0x4d42;        // 0x42 = "B" 0x4d = "M" 
    // Compute the size of the entire file. 
    hdr.bfSize = (DWORD) (sizeof(BITMAPFILEHEADER) + 
                 pbih->biSize + pbih->biClrUsed 
                 * sizeof(RGBQUAD) + pbih->biSizeImage); 
    hdr.bfReserved1 = 0; 
    hdr.bfReserved2 = 0; 

    // Compute the offset to the array of color indices. 
    hdr.bfOffBits = (DWORD) sizeof(BITMAPFILEHEADER) + 
                    pbih->biSize + pbih->biClrUsed 
                    * sizeof (RGBQUAD); 

    // Copy the BITMAPFILEHEADER into the .BMP file. 
    if (!WriteFile(hf, (LPVOID) &hdr, sizeof(BITMAPFILEHEADER), 
        (LPDWORD) &dwTmp,  NULL)) 
    {
       errhandler("WriteFile", hwnd);
	   goto cleanup;
    }

    // Copy the BITMAPINFOHEADER and RGBQUAD array into the file. 
    if (!WriteFile(hf, (LPVOID) pbih, sizeof(BITMAPINFOHEADER) 
                  + pbih->biClrUsed * sizeof (RGBQUAD), 
                  (LPDWORD) &dwTmp, ( NULL))) 
	{
        errhandler("WriteFile", hwnd); 
		goto cleanup;
	}

    // Copy the array of color indices into the .BMP file. 
    dwTotal = cb = pbih->biSizeImage; 
    hp = lpBits; 
    if (!WriteFile(hf, (LPSTR) hp, (int) cb, (LPDWORD) &dwTmp,NULL)) 
	{
           errhandler("WriteFile", hwnd); 
		   goto cleanup;
	}
    // Close the .BMP file. 
     if (!CloseHandle(hf)) 
	 {
           errhandler("CloseHandle", hwnd); 
		   goto cleanup;
	 }

	bSuccess = TRUE;
cleanup:
    // Free memory.
    if (lpBits != NULL)
		GlobalFree((HGLOBAL)lpBits);

	if (pbi != NULL)
		GlobalFree(pbi);

	return bSuccess;
}




STDMETHODIMP CSnapper::QueryStatus ( 
    /* [unique][in] */ const GUID __RPC_FAR *pguidCmdGroup,
    /* [in] */ ULONG cCmds,
    /* [out][in][size_is] */ OLECMD __RPC_FAR prgCmds[  ],
	/* [unique][out][in] */ OLECMDTEXT __RPC_FAR *pCmdText)
{
	for (ULONG i = 0; i < cCmds; i++)
		prgCmds[i].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED ;
	return S_OK;
}



#define MALFORMED_XML_ERROR 0x80041000 
#define INVALID_XML_ERROR 0x80041001 
#define ERROR_CREATING_SECTION 0x80041002 
#define ERROR_OPENING_SECTION 0x80041003 
#define SECTION_DOES_NOT_EXIST_ERROR 0x80041004 
#define PAGE_DOES_NOT_EXIST_ERROR 0x80041005 
#define FILE_DOES_NOT_EXIST_ERROR 0x80041006 
#define ERROR_INSERTING_IMAGE 0x80041007 
#define ERROR_INSERTING_INK 0x80041008 
#define ERROR_INSERTING_HTML 0x80041009 
#define ERROR_NAVIGATING_TO_PAGE 0x8004100a

void CSnapper::OneNoteErrorMsg(CAtlString& strErr, _com_error& e)
{

	switch (e.Error())
	{

		case MALFORMED_XML_ERROR:
			strErr = TEXT("Malformed import XML data");
			break;
		case INVALID_XML_ERROR:
			strErr = TEXT("Invalid import XML data");
			break;
		case ERROR_CREATING_SECTION:
			strErr = TEXT("Error creating OneNote section");
			break;
		case ERROR_OPENING_SECTION:
			strErr = TEXT("Error opening OneNote section.  If this is an encrypted section then you must first open the section in OneNote.");
			break;
		case SECTION_DOES_NOT_EXIST_ERROR:
			strErr = TEXT("OneNote section does not exist");
			break;
		case PAGE_DOES_NOT_EXIST_ERROR:
			strErr = TEXT("OneNote page does not exist");
			break;
		case FILE_DOES_NOT_EXIST_ERROR:
			strErr = TEXT("OneNote file does not exist");
			break;
		case ERROR_INSERTING_IMAGE:
			strErr = TEXT("OneNote error inserting image");
			break;
		case ERROR_INSERTING_INK:
			strErr = TEXT("OneNote error inserting ink");
			break;
		case ERROR_INSERTING_HTML:
			strErr = TEXT("OneNote error inserting HTML");
			break;
		case ERROR_NAVIGATING_TO_PAGE:
			strErr = TEXT("OneNote error navigating to page");
			break;
		default:
			strErr = e.ErrorMessage();
	}

}

