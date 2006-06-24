#include "StdAfx.h"
#include "SnapperXml.h"
#include "AlertNewSection.h"

#define THROW_ON_ERROR(fcn)	{ HRESULT hr; hr = fcn; if FAILED(hr) _com_issue_error(hr); }

CSnapperXml::CSnapperXml(void)
{
	Clear();

	HRESULT hr = m_spDom.CreateInstance(MSXML::CLSID_DOMDocument);
	if (FAILED(hr))
	{
		throw _T("problem instantiating MSXML");
	}
}

CSnapperXml::~CSnapperXml(void)
{
}

BOOL CSnapperXml::Load(LPCTSTR lpszFilename)
{

	// if the file doesn't exist then just return.
	if (INVALID_FILE_ATTRIBUTES == GetFileAttributes(lpszFilename))
		return FALSE;

	if (VARIANT_TRUE != m_spDom->load(CComVariant(lpszFilename)))
	{
		// there was an error.  We pretty-much know it isn't a FILE NOT FOUND because we've already checked for it.
		// if we want to be a bit more careful then we could check if it's 0x800c005 (access denied?) or whatever.
		CString strMsg = m_spDom->parseError->reason;
		CString strErr;
		strErr.Format(_T("error loading config file %s: %s [%x]"), lpszFilename, strMsg, m_spDom->parseError->errorCode);
		throw strErr;
	}

	return TRUE;
}

// Save to XML file.  Returns TRUE on success, throws an exception on failure.
void CSnapperXml::Save(LPCTSTR lpszFilename)
{
	m_spDom->save(CComVariant(lpszFilename));
}

void CSnapperXml::SetOption(LPCTSTR lpszOption, LPCTSTR lpszVal)
{
	MSXML::IXMLDOMElementPtr	spOptions = GetOptions();
	MSXML::IXMLDOMElementPtr	spOption;
	if (FindNamedNode(CString(lpszOption), spOptions, spOption))
	{	
		// option exists, so just re-set its value
		MSXML::IXMLDOMNodeListPtr spVals = spOption->getElementsByTagName(_T("val"));
		MSXML::IXMLDOMElementPtr spVal = spVals->item[0];
		spVal->nodeTypedValue = lpszVal;
	}	
	else 
	{
		// option doesn't yet exist, so create it
		spOption = m_spDom->createElement(_T("option"));
		MSXML::IXMLDOMElementPtr spName = m_spDom->createElement(_T("name"));
		spName->nodeTypedValue = lpszOption;
		spOption->appendChild(spName);
		MSXML::IXMLDOMElementPtr spVal = m_spDom->createElement(_T("val"));
		spVal->nodeTypedValue = lpszVal;
		spOption->appendChild(spVal);

		spOptions->appendChild(spOption);
	}

}
BOOL CSnapperXml::GetBoolOption(LPCTSTR strOpt, BOOL bDefault)
{
	LPTSTR	lpszDef = _T("FALSE");
	if (bDefault)
		lpszDef = _T("TRUE");

	CString strVal = GetOption(strOpt, lpszDef);
	return (0 == strVal.CompareNoCase(_T("TRUE")));
}
void CSnapperXml::SetBoolOption(LPCTSTR strOpt, BOOL bVal)	
{
	LPTSTR lpszVal = bVal ? _T("TRUE") : _T("FALSE");
	SetOption(strOpt, lpszVal);
}

CString CSnapperXml::GetOption(LPCTSTR lpszOption, LPCTSTR lpszDefault)
{
	MSXML::IXMLDOMElementPtr	spOptions = GetOptions();
	MSXML::IXMLDOMElementPtr	spOption;
	if (FindNamedNode(CString(lpszOption), spOptions, spOption))
	{	
		MSXML::IXMLDOMNodeListPtr spVals = spOption->getElementsByTagName(_T("val"));
		MSXML::IXMLDOMElementPtr spVal = spVals->item[0];
		return spVal->nodeTypedValue;
	}	
	else 
		return CString(lpszDefault);
}


long	CSnapperXml::GetLongOption(LPCTSTR strOpt, long lDefault)
{
	CString strDefault;
	strDefault.Format(L"%d", lDefault);
	CString strVal = GetOption(strOpt, strDefault);
	return _ttol( strVal );
}
void	CSnapperXml::SetLongOption(LPCTSTR strOpt, long lVal)
{
	CString strVal;
	strVal.Format(L"%d", lVal);
	SetOption(strOpt, strVal);
}


float	CSnapperXml::GetFloatOption(LPCTSTR strOpt, float fDefault)
{
	CString strDefault;
	strDefault.Format(L"%.2f", fDefault);
	CString strVal = GetOption(strOpt, strDefault);
	return (float) _tstof( strVal );
}
void	CSnapperXml::SetFloatOption(LPCTSTR strOpt, float fVal)
{
	CString strVal;
	strVal.Format(L"%.2f", fVal);
	SetOption(strOpt, strVal);
}
void CSnapperXml::GetTargets(CStringArray& rgTargets, LPCTSTR lpszHotlist)
{

	rgTargets.RemoveAll();

	MSXML::IXMLDOMElementPtr spTargets;
	if (GetTargets(CString(lpszHotlist), spTargets))
	{
		GetChildNames(rgTargets, spTargets);
	}

}
void CSnapperXml::GetChildNames(CStringArray& rgNames, MSXML::IXMLDOMElementPtr spRoot, BOOL bIncludeHidden)
{	
	rgNames.RemoveAll();

	MSXML::IXMLDOMNodeListPtr spNodes = spRoot->childNodes;

	for (int i = 0; i < spNodes->length; i++)
	{
		MSXML::IXMLDOMElementPtr spElem = spNodes->item[i];

		// if we're not including hidden children and the child is marked "hide" then skip it...
		_variant_t varHide = spElem->getAttribute(_bstr_t(L"hide"));
		if (bIncludeHidden
			|| (varHide.vt == VT_NULL) 
			|| (CString(L"true").CompareNoCase(varHide.bstrVal) != 0)
			)
		{
			MSXML::IXMLDOMElementPtr spName = spElem->getElementsByTagName(_T("name"))->item[0];
			rgNames.Add(spName->nodeTypedValue);
		}
	}

}

/** Mark all "modifiable" hotlists & and their targets for removal
 *
 *	@param bIgnoreNetworkResources			if TRUE then ignore any hotlist/section
 *											marked as a "network resource".  
 *
 *  @remarks
 *		bMarkRemoveNetworkResources is useful to perform "fast updates" of the 
 *		hotlists.  To do a "fast update", ignore all network resources when
 *		removing elements and when iterating through the notebook.
 */
void CSnapperXml::MarkRemoveHotlists(BOOL bIgnoreNetworkResources /* = FALSE */)
{
	MSXML::IXMLDOMElementPtr	spHotlists = GetHotlists();
	MSXML::IXMLDOMElementPtr	spHotlist = spHotlists->GetfirstChild();

	// iterate through all the hotlists...
	while (spHotlist != NULL)
	{
		// is the hotlist marked "nomodify", or is this a network resource and we're 
		// ignoring network resources?
		if (GetBoolAttrib(spHotlist, L"nomodify") == FALSE && 
			(bIgnoreNetworkResources == FALSE || !IsNetworkResource(spHotlist))
			)
		{
			// mark hotlist for removal
			MarkRemoveElement(spHotlist);
			// mark all its targets, too...
			MSXML::IXMLDOMElementPtr spTargets = GetNamedChild(spHotlist, L"targets");
			if (spTargets != NULL)
			{
				MSXML::IXMLDOMElementPtr spTarget = spTargets->GetfirstChild();
				while (spTarget != NULL)
				{
					if (GetBoolAttrib(spTarget, L"nomodify") == FALSE && 
						(bIgnoreNetworkResources == FALSE || !IsNetworkResource(spTarget))
						)
					{
						MarkRemoveElement(spTarget);					
					}
					spTarget = spTarget->GetnextSibling();
				};
			}
		}
		spHotlist = spHotlist->GetnextSibling();
	}
}
BOOL CSnapperXml::GetBoolAttrib(MSXML::IXMLDOMElementPtr& spElem, LPCTSTR lpszAttrib)
{
	BOOL bAttrib = FALSE;
	_variant_t varNoModify = spElem->getAttribute(_bstr_t(lpszAttrib));
	if ((varNoModify.vt != VT_NULL) && (CString(L"true").CompareNoCase(varNoModify.bstrVal) == 0))
		bAttrib = TRUE;
	return bAttrib;
}
BOOL IsInStringArray(CString& str, CStringArray& rgstr)
{
	for (int i = 0; i < rgstr.GetCount(); i++)
	{
		if (str.Compare(rgstr[i]) == 0)
			return TRUE;
	}
	return FALSE;
}
/** Reconcile hotlist/target "current" settings.
 *
 *	This function iterates through the hotlists and verifies its 
 *  "current" target is still in the hotlist (e.g.,, the 
 *	target hasn't been deleted).  If it's not in the hotlist
 *  then this function sets "current" to the first target in the hotlist.
 *
 *	It also verifies the current hotlist is in the list 
 *  of hotlists.  If not, it sets the current hotlist to the first hotlist.
 *
 *	@note: ignores @e nomodify
 *  
 */
void CSnapperXml::ReconcileCurrents()
{

	// make sure the current hotlist is in the list of hotlists
	CStringArray rgHotlists;
	GetHotlists(rgHotlists);
	if (!IsInStringArray(GetCurrentHotlist(), rgHotlists))
	{
		// oops - current hotlist not in list.  Set to the first hotlist.
		SetCurrentHotlist(rgHotlists[0]);
	}

	// verify each hotlist's current target is in the hotlist.
	for (int i = 0; i < rgHotlists.GetCount(); i++)
	{
		// make sure current target is in list of targets
		CStringArray rgTargets;
		CString& strHotlist = rgHotlists[i];
		GetTargets(rgTargets, strHotlist);
		if (!IsInStringArray( GetCurrentTarget(rgHotlists[i]), rgTargets))
		{	
			// oops - target not in list... Just assign current 
			// to the first item in the list, if any exist.
			// ASSERT(rgTargets.GetCount() > 0);
			if (rgTargets.GetCount() > 0)
				SetCurrentTarget(strHotlist, rgTargets[0]);
		}
	}

}



/** Remove hotlists/targets marked for removal
 *
 *  Hotlists are marked for removal by adding a "remove" attribute to its 
 *  corresponding element.
 *
 *  The "modifiable" field is ignored.
 */
void CSnapperXml::RemoveMarkedHotlists()
{
	MSXML::IXMLDOMElementPtr	spHotlists = GetHotlists();
	MSXML::IXMLDOMElementPtr	spHotlist = spHotlists->GetfirstChild();

	// iterate through all the hotlists...
	while (spHotlist != NULL)
	{
		MSXML::IXMLDOMElementPtr spNext = spHotlist->GetnextSibling();
		// is the hotlist marked for delete?
		BOOL bRemoveHotlist = FALSE;
		if (IsMarkedForRemoval(spHotlist))
		{
			bRemoveHotlist = TRUE;
		}
		else
		{
			// not marked for removal - but a target may be... so let's take a look...
			MSXML::IXMLDOMElementPtr spTargets = GetNamedChild(spHotlist, L"targets");
			if (spTargets != NULL)
			{
				MSXML::IXMLDOMElementPtr spTarget = spTargets->GetfirstChild();
				while (spTarget != NULL)
				{
					MSXML::IXMLDOMElementPtr spNextTarget = spTarget->GetnextSibling();
					if (IsMarkedForRemoval(spTarget))
					{
						spTargets->removeChild(spTarget);
					}
					spTarget = spNextTarget;
				};
				// we're going to keep empty folders in the list.  This shouldn't be a common 
				// occurence, but users may want thebe disturbed if their folder list doesn't have all
				// their folders. 
				// remove the hotlist if it doesn't have any targets
				//if (spTargets->GetfirstChild() == NULL)
				//	bRemoveHotlist = TRUE;
			}
			else
			{
				// this should never happen, but remove the hotlist if there's no target list
				ASSERT(0); 
				bRemoveHotlist = TRUE;
			}

		}
		if (bRemoveHotlist)
		{
			// I'm assuming removing the node doesn't muck w/ the order such that this is a reasonable
			// method for iterating through the elements...
			spHotlists->removeChild(spHotlist);
		}
		spHotlist = spNext;
	};
}

void CSnapperXml::RemoveHotlists()
{
	// remove the hotlists element, and create & attach a new one... 
	// not elegant, but good enough for now.
	MarkRemoveHotlists(FALSE);
	RemoveMarkedHotlists();
}
MSXML::IXMLDOMElementPtr CSnapperXml::GetRoot()
{
	return m_spDom->getElementsByTagName(_T("OneSnap"))->item[0];
}
MSXML::IXMLDOMElementPtr CSnapperXml::GetOptions()
{
	return m_spDom->getElementsByTagName(_T("options"))->item[0];
}
MSXML::IXMLDOMElementPtr CSnapperXml::GetHotlists()
{
	return m_spDom->getElementsByTagName(_T("hotlists"))->item[0];
}
void CSnapperXml::GetHotlists(CStringArray& rgHotlists)
{
	GetChildNames(rgHotlists, GetHotlists());//m_spHotlists);
}
// clear out the config, starting over w/ just an initial "shell" config.
void CSnapperXml::Clear(void)
{

	// remove all references to current DOM object and create a new DOM object.
	// old DOM object should nicely clean itself up.
	HRESULT hr = m_spDom.CreateInstance(MSXML::CLSID_DOMDocument);
	if (FAILED(hr))
	{
		throw _T("problem instantiating MSXML");
	}

	// we use XPath for some of our patter recognition..
	//m_spDom->setProperty(_T("SelectionLanguage"), _T("XPath"));


	// set output file to unicode.
	MSXML::IXMLDOMProcessingInstructionPtr spPi = m_spDom->createProcessingInstruction(_T("xml"), _T("version='1.0' encoding='UTF-16'"));
	_variant_t vNullVal;
	vNullVal.vt = VT_NULL;
	m_spDom->insertBefore(spPi, vNullVal);

	// create the "shell" nodes that make up the config file.

	MSXML::IXMLDOMElementPtr spRoot, spHotlists, spOptions;
	spRoot = m_spDom->createElement(_T("OneSnap"));
	spHotlists = m_spDom->createElement(_T("hotlists"));
	spRoot->appendChild(spHotlists);
	spOptions = m_spDom->createElement(_T("options"));
	spRoot->appendChild(spOptions);
	m_spDom->appendChild(spRoot);

}

BOOL CSnapperXml::FindNamedNode(CString& strName, MSXML::IXMLDOMElementPtr& spRootNode, MSXML::IXMLDOMElementPtr& spFoundNode)
{
	
	MSXML::IXMLDOMNodeListPtr spNodes = spRootNode->childNodes;
	BOOL bFound = FALSE;
	for (int i = 0; i < spNodes->length; i++)
	{
		MSXML::IXMLDOMElementPtr spNode = spNodes->item[i];
		if (MatchesName(spNode, strName))
		{
			spFoundNode = spNode;
			bFound = TRUE;
			break;
		}

	}
	return bFound;
}
BOOL CSnapperXml::FindHotlist(CString& strHotlist, MSXML::IXMLDOMElementPtr& spHotlist)
{
	return FindNamedNode(strHotlist, GetHotlists(), spHotlist);
}

BOOL CSnapperXml::MatchesName(MSXML::IXMLDOMElementPtr& spNode, CString& strName)
{
	MSXML::IXMLDOMElementPtr spName = spNode->getElementsByTagName(_T("name"))->item[0];
	if (strName == spName->nodeTypedValue)
	{
		return TRUE;
	}
	
	return FALSE;
}
BOOL CSnapperXml::FindTarget(CString& strHotlist, CString &strTarget, MSXML::IXMLDOMElementPtr& spTarget)
{

	BOOL bFound = FALSE;
	MSXML::IXMLDOMElementPtr spTargets;
	if (GetTargets(strHotlist, spTargets))
	{
		bFound = FindNamedNode(strTarget, spTargets, spTarget);
	}

	return bFound;
}
void CSnapperXml::AddTarget(LPCTSTR lpszHotlist, LPCTSTR lpszTarget, LPCTSTR lpszPath, BOOL bIsNetworkPath)
{

	MSXML::IXMLDOMElementPtr	spTarget;
	if (FindTarget(CString(lpszHotlist), CString(lpszTarget), spTarget))
	{
		// hotlist is already on the list - probably marked for removal.
		// so just unmark it for removal
		UnMarkRemoveElement(spTarget);
		SetNetworkResource(spTarget, bIsNetworkPath);
	}
	else 
		CreateTarget(lpszHotlist, lpszTarget, lpszPath, bIsNetworkPath);
}

void CSnapperXml::CreateTarget(LPCTSTR lpszHotlist, LPCTSTR lpszTarget, LPCTSTR lpszPath, BOOL bIsNetworkPath)
{
	MSXML::IXMLDOMElementPtr spTargets;
	if (GetTargets(CString(lpszHotlist), spTargets))
	{
		MSXML::IXMLDOMElementPtr spTarget = m_spDom->createElement(_T("target"));
		MSXML::IXMLDOMElementPtr spName = m_spDom->createElement(_T("name"));
		spName->nodeTypedValue = lpszTarget;
		spTarget->appendChild(spName);
		MSXML::IXMLDOMElementPtr spPath = m_spDom->createElement(_T("path"));
		spPath->nodeTypedValue = lpszPath;
		spTarget->appendChild(spPath);
		SetNetworkResource(spTarget, bIsNetworkPath);

		spTargets->appendChild(spTarget);

	}
	else 
	{
		CString strErr;
		strErr.Format(_T("error finding hotlist: %s"), lpszHotlist);
		throw strErr;
	}

}

void CSnapperXml::AddHotlist(LPCTSTR lpszHotlist, LPCTSTR lpszDefaultPath, BOOL bIsNetworkPath)
{
	MSXML::IXMLDOMElementPtr	spHotlist;
	if (FindHotlist(CString(lpszHotlist), spHotlist))
	{
		// hotlist is already on the list - probably marked for removal.
		// so just unmark it for removal
		// NOTE: there's a chance a new hotlist was created w/ the same name as a previous
		//       hotlist.  We're not going to worry about that small corner-case.
		UnMarkRemoveElement(spHotlist);
		// just in case we'll reset the network resource flag...
		SetNetworkResource(spHotlist, bIsNetworkPath);
	}
	else 
		CreateHotlist(lpszHotlist, lpszDefaultPath, bIsNetworkPath);
}
// Add (empty) hotlist to config file
void CSnapperXml::CreateHotlist(LPCTSTR lpszHotlist, LPCTSTR lpszDefaultPath, BOOL bIsNetworkPath)
{

	MSXML::IXMLDOMElementPtr spHotlist = m_spDom->createElement(_T("hotlist"));
	MSXML::IXMLDOMElementPtr spName = m_spDom->createElement(_T("name"));
	spName->nodeTypedValue = lpszHotlist;
	spHotlist->appendChild(spName);
	MSXML::IXMLDOMElementPtr spPath = m_spDom->createElement(_T("default_path"));
	spPath->nodeTypedValue = lpszDefaultPath;
	spHotlist->appendChild(spPath);
	SetNetworkResource(spHotlist, bIsNetworkPath);

	MSXML::IXMLDOMElementPtr spTargets = m_spDom->createElement(_T("targets"));
	spHotlist->appendChild(spTargets);

	GetHotlists()->appendChild(spHotlist);

}

BOOL CSnapperXml::GetTargets(CString& strHotlist, MSXML::IXMLDOMElementPtr& spTargets)
{

	BOOL bFound = FALSE;
	MSXML::IXMLDOMElementPtr spHotlist;
	if (FindHotlist(strHotlist, spHotlist))
	{
		MSXML::IXMLDOMNodeListPtr spTargetsList = spHotlist->getElementsByTagName(_T("targets"));
		spTargets = spTargetsList->item[0];
		bFound = TRUE;
	}

	return bFound;
}

// set currently-selected hotlist & target.  If target is not 
// already in hotlist, then add it.
void CSnapperXml::SetCurrent(LPCTSTR lpszHotlist, LPCTSTR lpszTarget)
{
	SetCurrentHotlist(lpszHotlist);
	SetCurrentTarget(lpszHotlist, lpszTarget);
}

void CSnapperXml::SetCurrentTarget(LPCTSTR lpszHotlist, LPCTSTR lpszTarget)
{
	MSXML::IXMLDOMElementPtr spTargets;
	if (GetTargets(CString(lpszHotlist), spTargets))
	{
		MSXML::IXMLDOMAttributePtr spAttr = m_spDom->createAttribute(_T("current"));
		spAttr->nodeTypedValue = lpszTarget;
		spTargets->attributes->setNamedItem(spAttr);

	}
	else
	{
		CString strErr;
		strErr.Format(_T("problem finding hotlist %s in config file."), lpszHotlist);
		throw strErr;
	}
}

void CSnapperXml::SetCurrent(MSXML::IXMLDOMElementPtr& spElement, LPCTSTR lpszCurrent)
{
	MSXML::IXMLDOMAttributePtr spAttr = m_spDom->createAttribute(_T("current"));
	spAttr->nodeTypedValue = lpszCurrent;
	spElement->attributes->setNamedItem(spAttr);
}
CString CSnapperXml::GetCurrent(MSXML::IXMLDOMElementPtr& spElement)
{
	CString strCur = _T("");
	MSXML::IXMLDOMAttributePtr spAttr = spElement->attributes->getNamedItem(_T("current"));
	strCur = spAttr ? spAttr->nodeTypedValue : _T("");
	return strCur;
}
void CSnapperXml::SetCurrentHotlist(LPCTSTR lpszHotlist)
{
	SetCurrent(GetHotlists(), lpszHotlist);
}


CString CSnapperXml::GetCurrentTarget(LPCTSTR lpszHotlist)
{
	MSXML::IXMLDOMElementPtr spTargets;
	if (GetTargets(CString(lpszHotlist), spTargets))
	{
		return GetCurrent(spTargets);
	}
	else
	{
		CString strErr;
		strErr.Format(_T("problem finding hotlist %s in config file."), lpszHotlist);
		throw strErr;
	}
}

CString CSnapperXml::GetCurrentHotlist()
{
	return GetCurrent(GetHotlists());
}
// get path to the .ONE file associated w/ this hotlist & target.
CString CSnapperXml::GetPath(LPCTSTR lpszHotlist, LPCTSTR lpszTarget)
{

	CString strHotlist(lpszHotlist);
	CString strTarget(lpszTarget);
	MSXML::IXMLDOMElementPtr spTarget;
	if (FindTarget(strHotlist, strTarget, spTarget))
	{
		MSXML::IXMLDOMNodeListPtr spPaths = spTarget->getElementsByTagName(_T("path"));
		MSXML::IXMLDOMElementPtr spPath = spPaths->item[0];
		return spPath->nodeTypedValue;
	}
	else 
	{
		CString strPath;
		if (QueryTargetPath(strHotlist,strTarget, strPath))
		{
			AddTarget(strHotlist, lpszTarget, strPath, PathIsNetworkPath(strPath));
			return strPath;
		}
		else 
			throw L"can't find target";
	}

}
BOOL CSnapperXml::QueryTargetPath(CString strHotlist, CString strTarget, CString& strPath)
{
	// target doesn't exist, so create it using the hotlist's default path
	// TODO: move new target creation somewhere else

	BOOL bGotTarget = FALSE;
	BOOL bAddTarget = TRUE;
	if (!GetBoolOption(L"skip_newsection_alert", FALSE))
	{
		CAlertNewSection dlgAlert;
		if (IDOK == dlgAlert.DoModal())
		{
			SetBoolOption(L"skip_newsection_alert", dlgAlert.m_bAlwaysSkip);
		}
		else
		{
			bAddTarget = FALSE;
		}
	}
	
	if (bAddTarget)
	{
		MSXML::IXMLDOMElementPtr spHotlist;
		if (FindHotlist(strHotlist, spHotlist))
		{
			MSXML::IXMLDOMElementPtr spDefaultPath = GetNamedChild(spHotlist, L"default_path");
			if (spDefaultPath != NULL)
			{
				CString strDefSecPath = spDefaultPath->nodeTypedValue;
				strDefSecPath += L"\\";
				strDefSecPath += strTarget;
				CFileDialog	dlgSecPath(TRUE, TEXT("one"), strDefSecPath, OFN_CREATEPROMPT | OFN_HIDEREADONLY, TEXT("OneNote Sections (*.one)|*.one||"));
				if (IDOK == dlgSecPath.DoModal())
				{
					strPath = dlgSecPath.GetPathName();
					bGotTarget = TRUE;
				}
			}
		}
	}

	return bGotTarget;
}


// set max width option
void CSnapperXml::SetMaxWidth(float fMaxWidth)
{
}

// get max width option
float CSnapperXml::GetMaxWidth()
{
	return 10.5f;
}


// set "image set to background" option
void CSnapperXml::SetBitmapAsBackground(BOOL bBackground)
{
}

// get "image set to background" option
BOOL CSnapperXml::GetBitmapAsBackground()
{
	return TRUE;
}
void CSnapperXml::UnMarkRemoveElement(MSXML::IXMLDOMElementPtr spElement)
{
	ASSERT(spElement != NULL);
	if (spElement != NULL)
		spElement->removeAttribute(_bstr_t(L"remove"));
}

void CSnapperXml::MarkRemoveElement(MSXML::IXMLDOMElementPtr spElement)
{
	ASSERT(spElement != NULL);
	if (spElement != NULL)
		spElement->setAttribute(_bstr_t(L"remove"), L"true");
}
BOOL CSnapperXml::IsMarkedForRemoval(MSXML::IXMLDOMElementPtr spElement)
{
	BOOL bRemove = FALSE;
	ASSERT(spElement != NULL);
	if (spElement != NULL)
	{
		_variant_t varRemove = spElement->getAttribute(_bstr_t(L"remove"));
		if (varRemove.vt != VT_NULL)
			bRemove = TRUE;
	}
	return bRemove;
}

void CSnapperXml::SetNetworkResource(MSXML::IXMLDOMElementPtr spElement, BOOL bOnNetwork)
{
	ASSERT(spElement != NULL);
	if (spElement != NULL)
	{
		if (bOnNetwork)
			spElement->setAttribute(_bstr_t(L"network_resource"), L"true");
		else
			spElement->removeAttribute(_bstr_t(L"network_resource"));
	}

}
BOOL CSnapperXml::IsNetworkResource(MSXML::IXMLDOMElementPtr spElement)
{
	return GetBoolAttrib(spElement, L"network_resource");
}




