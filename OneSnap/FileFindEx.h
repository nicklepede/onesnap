/** @file	FileFindEx.h
 *  @brief	CFileFind wrapper that supports shortcuts
 *  @author Andrew Wheeler
 */
#pragma once
#include "afx.h"

/** Simple file/directory iterator that can follow shortcuts
  *
  * Loosely based on CFileFind.
  */
class CFileIterator
{
public:
	CFileIterator(LPCTSTR lpszDir, LPCTSTR lpszFilespec = _T("*.*"), BOOL bFollowShortcuts = TRUE, BOOL bFollowNetworkPaths = TRUE);

	/** Iterate to next file object
	 *	
	 *  @returns TRUE if a next file was found, otherwise FALSE
	 */
	BOOL Next(void);

	CString GetPath();
	CString GetTitle();
	CString GetName();
	BOOL	IsDir();
	BOOL	IsFile();
	BOOL	IsNetworkPath();

public:
	~CFileIterator(void);

private:
	HRESULT OpenLink(LPCTSTR	lpszPath, IShellLink** ppiLink) const;
	HRESULT GetShortcutInfo(LPCTSTR pszShortcutPath, LPTSTR pszTargetPath, int cchMaxPath, WIN32_FIND_DATA* psFd) const;
	CFileFind	m_ff;				// 
	BOOL m_bFollowShortcuts;		// if TRUE then FindNextFile should follow shortcuts
	BOOL m_bFollowNetworkPaths;		// if TRUE then FindNextFile should include/follow network paths
	BOOL m_bIsShortcut;				// if TRUE then use m_strPathShortcut instead of the m_ff's current file.
	CString	m_strPath;				// if m_bIsShortcut then this is the shortcut's target path.
	WIN32_FIND_DATA m_sFd;			// if m_bIsShortcut then this is the shortcut's file data.
	CString m_strFilespec;			// DOS-type file matching specification 
	BOOL	m_bFoundFile;			// TRUE when we're currently "pointing" at a file object
	BOOL	m_bNoMoreFiles;			// TRUE if there's no more files to iterate through
};
