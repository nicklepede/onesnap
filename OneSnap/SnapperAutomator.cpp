// SnapperAutomator.cpp : Implementation of CSnapperAutomator

#include "stdafx.h"
#include "SnapperAutomator.h"
#include "Snapper.h"

// CSnapperAutomator


STDMETHODIMP CSnapperAutomator::SnapIt(IUnknown* pUnknown)
{
//	AFX_MANAGE_STATE(AfxGetStaticModuleState());


	CComObject<CSnapper>	oSnapper;

	HRESULT hr = oSnapper.SetSite(pUnknown);

	if SUCCEEDED(hr)
	{

		hr = oSnapper.Exec();

	}


	return hr;
}
