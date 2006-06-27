#include "StdAfx.h"
#include "snapperconfig.h"
#include "FileFindEx.h"

#define ONESNAP_REGKEY_PATH		_T("Software\\OneSnap")
#define ONESNAP_REGVAL_FILEPATH	_T("Filepath")
#define ONESNAP_REGVAL_NOTEBOOKPATH	_T("NotebookPath")

// options and their default values
#define OPT_LIMITPAGEWIDTH	_T("limit_width")
#define DEF_LIMITPAGEWIDTH	FALSE

#define OPT_MAXWIDTH	_T("max_width_x100")
#define DEF_MAXWIDTH	12.0f

#define OPT_IMPORTASBACKGROUND	_T("import_as_background")
#define DEF_IMPORTASBACKGROUND	FALSE

#define OPT_NAVIGATETOPAGE		_T("navigate_to_page")
#define DEF_NAVIGATETOPAGE		FALSE

#define OPT_INCLUDESHAREDNOTEBOOKS		_T("include_shared_notebooks")
#define DEF_INCLUDESHAREDNOTEBOOKS		FALSE

#define OPT_AUTOSYNC	_T("auto_sync")
#define DEF_AUTOSYNC	TRUE

#define OPT_FOLDER_SEPARATOR	_T("folder_separator")
#define DEF_FOLDER_SEPARATOR	_T("\\")

#define OPT_MAX_DIR_DEPTH	_T("max_dir_depth")
#define DEF_MAX_DIR_DEPTH	10


// load a value from OneSnap's registry key.  If the value doesn't exist, 
// then return the default value *and* save the default value to the registry
// (this isn't totally necessary, but it's convenient for manual regval tweaking).

// GetFileTitle is defined to be GetFileTitleW, but GetFileTitle is the name of a member function we must call
// below.  Therefore we must undef GetFileTitle.
#undef GetFileTitle


CSnapperConfig::CSnapperConfig()
{
	if (!Load())
	{
		CreateConfigFile(m_strFilepath, m_strNotebookPath);
	}
	else
	{
		if (GetAutoSync())
			UpdateHotlists(FALSE);
	}
}

CSnapperConfig::~CSnapperConfig(void)
{
}

BOOL CSnapperConfig::Load(void)
{
	LoadRegSettings();
	return CSnapperXml::Load(m_strFilepath);
}
#define REGVAL_MAXLEN	512
CString CSnapperConfig::LoadFromRegistry(LPCTSTR lpszValName, LPCTSTR lpszDefault)
{
	CString	strVal;
	strVal.Preallocate(REGVAL_MAXLEN);
	strVal = lpszDefault;

	CRegKey rk;
	if(rk.Create(HKEY_CURRENT_USER, ONESNAP_REGKEY_PATH) == ERROR_SUCCESS)
	{
		ULONG nChars = strVal.GetAllocLength();
		if (ERROR_SUCCESS != rk.QueryStringValue(lpszValName, strVal.GetBuffer(), &nChars))
		{
			// wasn't successful at reading setting, so save the default settting.
			rk.SetStringValue(lpszValName, strVal);
		}
		strVal.ReleaseBuffer();
		rk.Close();
	}	

	return strVal;
}

void CSnapperConfig::AddCategoryEntry(LPCTSTR lpszPath, LPCTSTR lpszCategory, BOOL bIsNetworkPath, BOOL bFollowShortcuts, BOOL bFollowNetworkPaths)
{
	AddHotlist(lpszCategory, lpszPath, bIsNetworkPath);

	// find all .one files and add them as targets
	CFileIterator	fi(lpszPath, L"*.one", bFollowShortcuts, bFollowNetworkPaths);

	while (fi.Next())
	{
		CString strFilePath = fi.GetPath();
		// CFileIterator does 8.3 pattern matches, so unfortunately it catches the 
		// *.onetoc files when we filter for *.one files.  So let's filter them out
		if (strFilePath.Find(_T(".onetoc"), strFilePath.GetLength() - 8) != -1)
			continue;

		AddTarget(lpszCategory, fi.GetTitle(), fi.GetPath(), fi.IsNetworkPath());
	};
}

void CSnapperConfig::AddCategories(LPCTSTR lpszPath, LPCTSTR lpszName, BOOL bFollowShortcuts, BOOL bFollowNetworkPaths, BOOL bPrependNameToSubs, LPCTSTR pszSeparator, long nMaxDirDepth)
{
	// if we've recursed to our max recurse depth, then quietly bail out.  Otherwise, continue 
	// one w/ our max depth decremented... 
	if (nMaxDirDepth-- <= 0)
	{
		return;
	}
	AddCategoryEntry(lpszPath, lpszName, PathIsNetworkPath(lpszPath), bFollowShortcuts, bFollowNetworkPaths);

	CFileIterator ff(lpszPath, _T("*.*"), bFollowShortcuts, bFollowNetworkPaths);
	while (ff.Next())
	{
		if (ff.IsDir())
		{
			CString strDirName = ff.GetTitle();
			if (bPrependNameToSubs)
				strDirName = CString(lpszName) + pszSeparator + strDirName;

			AddCategories(ff.GetPath(), strDirName, bFollowShortcuts, bFollowNetworkPaths, TRUE, pszSeparator, nMaxDirDepth);
		}
	};

}

void CSnapperConfig::UpdateHotlists(BOOL bFollowNetworkPaths)
{
	// save the current hotlist & targets so we can reset them later
//	CString strHotlist = GetCurrentHotlist();
//	CString strTarget = GetCurrentTarget(strHotlist);
	MarkRemoveHotlists(!bFollowNetworkPaths);
	AddCategories(m_strNotebookPath, _T("My Notebook"),GetIncludeSharedNotebooks(), bFollowNetworkPaths, FALSE, GetFolderSeparator(), GetMaxDirDepth() );
	// if the current hotlist/target no longer exist then 

	RemoveMarkedHotlists();
	Save();
	ReconcileCurrents();
}
void CSnapperConfig::CreateConfigFile(LPCTSTR lpszFilepath, LPCTSTR lpszNotebookPath)
{
	Clear();
	AddCategories(lpszNotebookPath, _T("My Notebook"), GetIncludeSharedNotebooks(), TRUE, TRUE, GetFolderSeparator(), GetMaxDirDepth());
	ReconcileCurrents();
	InitOptions();
	Save();
}

void CSnapperConfig::Save(void)
{
	SaveRegSettings();
	CSnapperXml::Save(m_strFilepath);
}


// save value to OneSnap's registry key.  Don't bother throwing if 
// something goes wrong...
BOOL CSnapperConfig::SaveToRegistry(LPCTSTR lpszValName, LPCTSTR lpszVal)
{
	long ns;
	CRegKey rk;
	ns = rk.Create(HKEY_CURRENT_USER, ONESNAP_REGKEY_PATH);
	if(ns == ERROR_SUCCESS)
	{
		ns = rk.SetStringValue(lpszValName, lpszVal);
		rk.Close();
	}	

	return (ns == ERROR_SUCCESS);
}


// load XML filepath and notebook filepath from registry, or
// set to default values.  
void CSnapperConfig::LoadRegSettings()
{
	// preset settings to default values
	m_strNotebookPath = LoadFromRegistry(ONESNAP_REGVAL_NOTEBOOKPATH, GetDefaultNotebookPath());
	// verify path points to a directory. If it doesn't then just ask the user to select it.
	if (!PathIsDirectory(m_strNotebookPath))
	{
		MessageBox(NULL, _T("OneSnap couldn't find your OneNote notebook.  After clicking OK you'll get a directory browser.  Please use it to find and select your notebook."), _T("Couldn't find notebook"), MB_OK);
		BrowseForNotebook(m_strNotebookPath.GetBuffer());
	}
	m_strFilepath = LoadFromRegistry(ONESNAP_REGVAL_FILEPATH, GetDefaultFilepath(m_strNotebookPath));
}

void CSnapperConfig::SaveRegSettings()
{
	SaveToRegistry(ONESNAP_REGVAL_NOTEBOOKPATH, m_strNotebookPath);
	SaveToRegistry(ONESNAP_REGVAL_FILEPATH, m_strFilepath);
}

// get the default path to the config file.   Default path is <NOTEBOOK PATH>\OneSnap.xml.
CString CSnapperConfig::GetDefaultFilepath(CString& strNotebookPath)
{
	CString strFilePath;
	PathCombine(strFilePath.GetBufferSetLength(MAX_PATH), strNotebookPath, _T("OneSnap.xml"));
	strFilePath.ReleaseBuffer();
	return strFilePath;
}

// get the default path to "My Notebook".
CString CSnapperConfig::GetDefaultNotebookPath()
{
	CString strPath(_T(""), MAX_PATH);
	
	// the relative path from My Documents to the root notebook is stored in
	// a regkey.
	CRegKey rkSave;
	if (ERROR_SUCCESS == rkSave.Open(HKEY_CURRENT_USER, _T("Software\\Microsoft\\Office\\11.0\\OneNote\\Options\\Save"), KEY_READ))
	{
		CString strRelPath;
		ULONG cch = MAX_PATH;
		if (ERROR_SUCCESS == rkSave.QueryStringValue(_T("My Notebook path"), strRelPath.GetBufferSetLength(MAX_PATH), &cch))
		{
			strRelPath.ReleaseBuffer();

			CString strPathMyDocs;
			if (S_OK == SHGetFolderPath(NULL, CSIDL_PERSONAL, NULL, 0,  strPathMyDocs.GetBufferSetLength(MAX_PATH)))
			{
				strPathMyDocs.ReleaseBuffer();
				PathCombine(strPath.GetBuffer(), strPathMyDocs, strRelPath);
				strPath.ReleaseBuffer();
			}
		}
	}

	return strPath;
}

bool CSnapperConfig::BrowseForNotebook(TCHAR* pszNotebookPath)
{
	bool bGotit = false;
	BROWSEINFO sBi = { 0 };
	sBi.lpszTitle = _T("Select your notebook directory");
	/*
    sBi.lpszTitle = TEXT("Select Notebook(s)");
	sBi.hwndOwner = m_hWndTop;
	sBi.iImage = 
	*/
    LPITEMIDLIST pIdl = SHBrowseForFolder ( &sBi );
    if ( pIdl != NULL )
    {
        // get the name of the folder
        if ( SHGetPathFromIDList ( pIdl, pszNotebookPath ) )
        {
			bGotit = true;
        }

        // free memory 
        IMalloc * piMalloc = NULL;
        if ( SUCCEEDED( SHGetMalloc ( &piMalloc )) )
        {
            piMalloc->Free ( pIdl );
            piMalloc->Release ( );
        }

    }
	
	return bGotit;
	
}




#if 0
		string notebookPath = "My Notebook";

			// The Notebook Path is stored in the Save options in the registry:
			string saveKey = "Software\\Microsoft\\Office\\11.0\\OneNote\\Options\\Save";
			using (RegistryKey saveOptions = Registry.CurrentUser.OpenSubKey(saveKey))
			{
				if (saveOptions != null)
					notebookPath = saveOptions.GetValue("My Notebook path", notebookPath).ToString();				
			}

			// This path is relative to the user's My Documents folder
			string documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
			return Path.Combine(documentsFolder, notebookPath);
#endif


void CSnapperConfig::SetConfigFilePath(LPCTSTR lpszFilepath)
{
	m_strFilepath = lpszFilepath;
	// load the config file if it exists, otherwise just save to the new filepath.
	if (!CSnapperXml::Load(m_strFilepath))
		CSnapperXml::Save(m_strFilepath);

	SaveRegSettings();
		
}

void CSnapperConfig::SetNotebookPath(LPCTSTR lpszNotebookPath)
{
	m_strNotebookPath = lpszNotebookPath;
}

// returns the path to the OneNote notebook
CString CSnapperConfig::GetNotebookPath(void)
{
	return m_strNotebookPath;
}

CString CSnapperConfig::GetConfigFilePath(void)
{
	return m_strFilepath;
}

CString CSnapperConfig::GetFolderSeparator()
{
	CString strSep = GetOption(OPT_FOLDER_SEPARATOR, DEF_FOLDER_SEPARATOR);
	// MSXML seems to clip leading and trailing whitespace.  
	// HACK: This is the one field where whitespace matters, so let's (sloppily) allow this field to use enclosing quotes 
	strSep.Replace(_T("\""), _T(""));
	return strSep;
}

void CSnapperConfig::SetFolderSeparator(LPCTSTR pszSeparator)
{
	SetOption(OPT_FOLDER_SEPARATOR, pszSeparator);
}
void CSnapperConfig::InitOptions()
{
	SetImportAsBackground(DEF_IMPORTASBACKGROUND);
	SetNavigateToPage(DEF_NAVIGATETOPAGE);
	SetLimitPageWidth(DEF_LIMITPAGEWIDTH);
	SetMaxPageWidth(DEF_MAXWIDTH);
	SetFolderSeparator(DEF_FOLDER_SEPARATOR);
	SetMaxDirDepth(DEF_MAX_DIR_DEPTH);
}
BOOL CSnapperConfig::GetImportAsBackground(void)
{
	return GetBoolOption(OPT_IMPORTASBACKGROUND, DEF_IMPORTASBACKGROUND);
}

BOOL CSnapperConfig::GetNavigateToPage(void)
{
	return GetBoolOption(OPT_NAVIGATETOPAGE, DEF_NAVIGATETOPAGE);
}

BOOL CSnapperConfig::GetIncludeSharedNotebooks(void)
{
	return GetBoolOption(OPT_INCLUDESHAREDNOTEBOOKS, DEF_INCLUDESHAREDNOTEBOOKS);
}

BOOL CSnapperConfig::GetLimitPageWidth(void)
{
	return GetBoolOption(OPT_LIMITPAGEWIDTH, DEF_LIMITPAGEWIDTH);
}

long CSnapperConfig::GetMaxDirDepth(void)
{
	return GetLongOption(OPT_MAX_DIR_DEPTH, DEF_MAX_DIR_DEPTH);
}

void CSnapperConfig::SetMaxDirDepth(long lDepth)
{
	return SetLongOption(OPT_MAX_DIR_DEPTH, lDepth);
}

float CSnapperConfig::GetMaxPageWidth(void)
{
	return GetFloatOption(OPT_MAXWIDTH, DEF_MAXWIDTH);
}

void CSnapperConfig::SetImportAsBackground(BOOL bImportAsBackground)
{
	SetBoolOption(OPT_IMPORTASBACKGROUND, bImportAsBackground);
}

void CSnapperConfig::SetNavigateToPage(BOOL bNavigateToPage)
{
	SetBoolOption(OPT_NAVIGATETOPAGE, bNavigateToPage);
}

void CSnapperConfig::SetIncludeSharedNotebooks(BOOL bIncludeSharedNotebooks)
{
	SetBoolOption(OPT_INCLUDESHAREDNOTEBOOKS, bIncludeSharedNotebooks);
}

void CSnapperConfig::SetLimitPageWidth(BOOL bLimitPageWidth)
{
	SetBoolOption(OPT_LIMITPAGEWIDTH, bLimitPageWidth);
}

void CSnapperConfig::SetMaxPageWidth(float fMaxWidth)
{
	SetFloatOption(OPT_MAXWIDTH, fMaxWidth);
}


// If this option is true, then we should rescan notebook each time user runs OneSnap.
BOOL CSnapperConfig::GetAutoSync(void)
{
	return GetBoolOption(OPT_AUTOSYNC, DEF_AUTOSYNC);
}
void CSnapperConfig::SetAutoSync(BOOL bAutoSync)
{
	SetBoolOption(OPT_AUTOSYNC, bAutoSync);
}
