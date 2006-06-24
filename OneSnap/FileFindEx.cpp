#include "StdAfx.h"
#include "FileFindEx.h"

/** Max path length for shortcut targets */
#define MAX_SHORTCUT_PATH	(MAX_PATH*4)

/** Constructor
 *
 *	@param lpszDir			directory to iterate
 *	@param lpszFilespec		DOS-style filename extension to match
 *	@param bFollowShortcuts	If TRUE then will iterate all shortcut targets instead of the 
 *					        .lnk files, themselves
 *	@param bFollowNetworkShortcuts	if TRUE and bFollowShortcuts is TRUE then iterate 
 *									shortcut targets that point to remote files/folders.
 */
CFileIterator::CFileIterator(LPCTSTR lpszDir, LPCTSTR lpszFilespec, BOOL bFollowShortcuts, BOOL bFollowNetworkPaths)
{
	m_bFollowShortcuts = bFollowShortcuts;
	m_bFollowNetworkPaths = bFollowNetworkPaths;
	m_strFilespec = lpszFilespec;
	// create a path spec that matches all files.  Since we may need to follow shortcuts, 
	// we have to match all files, and look for the file matches ourselves
	CString strDirAllFiles = lpszDir;
	strDirAllFiles += _T("\\*.*");

	// act as though there's no more files if we can't successfully initialize the CFileFind object.
	m_bNoMoreFiles = (0 == m_ff.FindFile(strDirAllFiles));
	m_bFoundFile = FALSE;
}

CFileIterator::~CFileIterator(void)
{
}

/** Iterate to next file object
 *	
 *  @returns TRUE if a next file was found, otherwise FALSE
 */
BOOL CFileIterator::Next(void)
{
	// just return if there's no more files or we weren't successfully initialized
	if (m_bNoMoreFiles) 
	{
		m_bFoundFile = FALSE;
		return FALSE;
	}
	BOOL bFound = FALSE;
	BOOL bMoreFiles;
	do 
	{
		bMoreFiles = (0 != m_ff.FindNextFile());
		if (m_ff.IsDots())
			continue;
		// if we know this isn't a shortcut, or we're not following shortcuts, then 
		// see if we have a filespec match
		if (!m_bFollowShortcuts
			|| m_ff.IsDirectory()
			|| !PathMatchSpec(m_ff.GetFileName(), _T("*.lnk"))
			)
		{
			if (m_bFollowNetworkPaths || !PathIsNetworkPath(m_ff.GetFilePath()))
			{
				if (PathMatchSpec(m_ff.GetFileName(), m_strFilespec))
				{
					bFound = TRUE;
					m_bIsShortcut = FALSE;
				}
			}
		}
		else
		{
			// this may be shortcut.   Try following it...
			WIN32_FIND_DATA fd;
			CString strTargetPath;
			if (SUCCEEDED(GetShortcutInfo(m_ff.GetFilePath(), strTargetPath.GetBuffer(MAX_SHORTCUT_PATH), MAX_SHORTCUT_PATH, &fd)))
			{
				// this is a shortcut.  If we're ignoring shortcuts to remote 
				// files/folders, then verify it's not a network path.
				if (m_bFollowNetworkPaths || !PathIsNetworkPath(strTargetPath))
				{
					// okay, try to match the filespec to it
					if (PathMatchSpec(fd.cFileName, m_strFilespec))
					{	
						bFound = TRUE;
						m_bIsShortcut = TRUE;
						m_strPath = strTargetPath;
						m_sFd = fd;
					}
				}

			}
			
			m_strPath.ReleaseBuffer();

		}

	} while (bFound == FALSE && bMoreFiles == TRUE);

	m_bFoundFile = bFound;
	m_bNoMoreFiles = !bMoreFiles;

	return bFound;
}
BOOL CFileIterator::IsNetworkPath()
{
	// return FALSE if nothing found
	if (!m_bFoundFile)
		return FALSE;
	CString strPath = GetPath();
	return PathIsNetworkPath(strPath);
}

BOOL CFileIterator::IsDir()
{
	// return an empty string if nothing has been found
	if (!m_bFoundFile)
		return FALSE;
	if (!m_bIsShortcut)
		// the found file isn't a shortcut, so get the info from the slaved file finder...
		return m_ff.IsDirectory();

	// we're currently on a shortcut, so get the filename from the cached file data
	return (m_sFd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

/**  Get shortcut info on current file
 *
 *	 @param[in]		pszPathShortcut	path to shortcut file
 *	 @param[out]	pszPathTarget	variable to be set to shortcut's target path
 *	 @param[in]		cchMaxPath		length of pszPathTarget buffer, in characters
 *	 @param[out]	psFd			shortcut's target info
 *
 *	 @returns successful @b HRESULT on success, else failing @b HRESULT on 
 *			  on error or if the current file is not a link.
 */
HRESULT CFileIterator::GetShortcutInfo(LPCTSTR pszPathShortcut, LPTSTR pszPathTarget, int cchMaxPath, WIN32_FIND_DATA* psFd) const
{
	// try opening the link...
	IShellLink* piLink;
	HRESULT hr = OpenLink(pszPathShortcut, &piLink);
	if (SUCCEEDED(hr))
	{
		// Get the path to the link target. 
	    hr = piLink->GetPath(pszPathTarget, cchMaxPath, psFd, SLGP_UNCPRIORITY);
		piLink->Release();
	}
	return hr;
}

/**  Open link object
 *
 *	 @param[in] lpszPath path to .lnk file
 *
 *   @returns pointer to @b IShellLink interface of link object on success,
 *		      else NULL.
 */
HRESULT CFileIterator::OpenLink(LPCTSTR	lpszPath, IShellLink** ppiLink) const
{
	HRESULT hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, 
                            IID_IShellLink, (LPVOID*)ppiLink);
	if (SUCCEEDED(hr))
	{
		BOOL bLoaded = FALSE;
		// try loading this file - it'll fail if it's not a .lnk, right?
		IPersistFile* piFile;
		hr = (*ppiLink)->QueryInterface(IID_IPersistFile, (void**) &piFile);
        if (SUCCEEDED(hr))
		{
			_bstr_t	bstrPath(lpszPath);
			hr = piFile->Load(bstrPath, 0);
			piFile->Release();
		}
		if (FAILED(hr))
		{
			// couldn't load this file, so return NULL
			(*ppiLink)->Release();
			*ppiLink = NULL;
		}

	}

	return hr;
}
CString CFileIterator::GetPath( )
{
	// return an empty string if nothing has been found
	if (!m_bFoundFile)
		return CString();
	// if we're currently on a shortcut then we've cached the target path
	if (m_bIsShortcut)
		return m_strPath;
	// not a shortcut...
	else 
		return m_ff.GetFilePath();
}
CString CFileIterator::GetName( )
{
	// return an empty string if nothing has been found
	if (!m_bFoundFile)
		return CString();
	if (!m_bIsShortcut)
		// the found file isn't a shortcut, so get the info from the slaved file finder...
		return m_ff.GetFileName();

	// we're currently on a shortcut, so get the filename from the cached file data
	CString strFileName = m_sFd.cFileName;
	return strFileName;
}
CString CFileIterator::GetTitle( )
{
	// return an empty string if nothing has been found
	if (!m_bFoundFile)
		return CString();

	// get the file's title.  If it's a shortcut then we still use the shortcut's 
	// title instead of the target file's title, as that's what the user sees
	// as their "exploring" the shortcut.
	return m_ff.GetFileTitle();

}
