// SnapperXml.h
//
// XML-file based config for OneSnap.
//
#pragma once

class CSnapperXml
{

public:
	CSnapperXml(void);
	~CSnapperXml(void);

	// load config file.  returns FALSE if file not found; throws an exception for other errors.
	BOOL Load(LPCTSTR lpszFilename);
	// Save to XML file.  Returns TRUE on success, FALSE on failure
	void Save(LPCTSTR lpszFilename);

	// remove everything from config.
 	void Clear(void);
	// remove all hotlists from the config
	void RemoveHotlists();

	// get/set option
	CString GetOption(LPCTSTR lpszOption, LPCTSTR lpszDefault);
	void	SetOption(LPCTSTR lpszOption, LPCTSTR lpszVal);
	BOOL	GetBoolOption(LPCTSTR strOpt, BOOL bDefault);
	void	SetBoolOption(LPCTSTR strOpt, BOOL bVal);
	float	GetFloatOption(LPCTSTR strOpt, float fDefault);
	void	SetFloatOption(LPCTSTR strOpt, float fVal);

	long	GetLongOption(LPCTSTR strOpt, long lDefault);
	void	SetLongOption(LPCTSTR strOpt, long lVal);

	// get specified hotlist's list of targets
	void GetTargets(CStringArray& rgTargets, LPCTSTR lpszHotlist);
	void GetHotlists(CStringArray& rgHotlists);

	void AddHotlist(LPCTSTR lpszHotlist, LPCTSTR lpszPath, BOOL bIsNetworkPath);
	void AddTarget(LPCTSTR lpszHotlist, LPCTSTR lpszTarget, LPCTSTR lpszPath, BOOL bIsNetworkPath);
	void CreateHotlist(LPCTSTR lpszHotlist, LPCTSTR lpszPath, BOOL bIsNetworkPath);
	void CreateTarget(LPCTSTR lpszHotlist, LPCTSTR lpszTarget, LPCTSTR lpszPath, BOOL bIsNetworkPath);
	
	// set currently-selected hotlist & target.  If target is not 
	// already in hotlist, then add it.
	void SetCurrent(LPCTSTR lpszHotlist, LPCTSTR lpszTarget);

	// set the current target of the specified hotlist. 
	// if target is not already in hotlist, then add it.
	// DOES NOT set the current hotlist.
	void SetCurrentTarget(LPCTSTR lpszHotlist, LPCTSTR lpszTarget);

	// set the current hotlist.
	void SetCurrentHotlist(LPCTSTR lpszHotlist);

	// get current defaults
	CString GetCurrentHotlist();
	CString GetCurrentTarget(LPCTSTR lpszHotlist);
	CString GetCurrent(CString& strHotlist, CString& strTarget);

	// get path to the .ONE file associated w/ this hotlist & target.
	CString GetPath(LPCTSTR lpszHotlist, LPCTSTR lpszTarget);

	void SetMaxWidth(float fMaxWidth);
	float GetMaxWidth();

	void SetBitmapAsBackground(BOOL bBackground);
	BOOL GetBitmapAsBackground();
protected:
	BOOL FindTarget(CString& strHotlist, CString& strTarget, MSXML::IXMLDOMElementPtr& spTarget);
	void MarkRemoveHotlists(BOOL bIgnoreNetworkResources);
	void RemoveMarkedHotlists();
	void ReconcileCurrents();

private:
	// find the hotlist node
	BOOL GetTargets(CString& strHotlist, MSXML::IXMLDOMElementPtr& spTargets);
	void GetChildNames(CStringArray& rgNames, MSXML::IXMLDOMElementPtr spElement, BOOL bIncludeHidden=FALSE);
	BOOL FindNamedNode(CString& strName, MSXML::IXMLDOMElementPtr& spRootNode, MSXML::IXMLDOMElementPtr& spFoundNode);
	BOOL FindHotlist(CString& strHotlist, MSXML::IXMLDOMElementPtr& spHotlist);
	BOOL MatchesName(MSXML::IXMLDOMElementPtr& spNode, CString& strName);
	MSXML::IXMLDOMElementPtr GetRoot();
	void SetCurrent(MSXML::IXMLDOMElementPtr& spElement, LPCTSTR lpszCurrent);
	CString GetCurrent(MSXML::IXMLDOMElementPtr& spElement);
	MSXML::IXMLDOMElementPtr CSnapperXml::GetHotlists();
	MSXML::IXMLDOMElementPtr GetOptions();

	void MarkRemoveElement(MSXML::IXMLDOMElementPtr spElement);
	void UnMarkRemoveElement(MSXML::IXMLDOMElementPtr spElement);
	BOOL IsMarkedForRemoval(MSXML::IXMLDOMElementPtr spElement);
	BOOL QueryTargetPath(CString strHotlist, CString strTarget, CString& strPath);
	MSXML::IXMLDOMElementPtr GetNamedChild(MSXML::IXMLDOMElementPtr& spElement, LPCTSTR lpszTagName)
	{
		MSXML::IXMLDOMNodeListPtr spVals = spElement->getElementsByTagName(lpszTagName);
		return spVals->item[0];
	};
	BOOL GetBoolAttrib(MSXML::IXMLDOMElementPtr& spElem, LPCTSTR lpszAttrib);
	BOOL IsNetworkResource(MSXML::IXMLDOMElementPtr spElement);
	void SetNetworkResource(MSXML::IXMLDOMElementPtr spElement, BOOL bOnNetwork);
	// MSXML stuff
	MSXML::IXMLDOMDocumentPtr	m_spDom;		// config file document object
};
