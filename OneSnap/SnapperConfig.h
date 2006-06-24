#pragma once

#include "SnapperXml.h"

class CSnapperConfig : public CSnapperXml
{
public:
	CSnapperConfig(void);
	~CSnapperConfig(void);
	BOOL Load(void);
	void Save(void);

	// returns the path to the OneNote notebook
	CString GetNotebookPath(void);
	CString GetConfigFilePath(void);
	void SetNotebookPath(LPCTSTR lpszNotebookPath);
	void SetConfigFilePath(LPCTSTR lpszFilepath);

	// update hotlists to reflect current notebook
	void UpdateHotlists(BOOL bFollowNetworkShortcuts);

	// option getters/setters
	BOOL GetImportAsBackground(void);
	BOOL GetNavigateToPage(void);
	BOOL GetLimitPageWidth(void);
	BOOL GetIncludeSharedNotebooks(void);
	float GetMaxPageWidth(void);
	CString GetFolderSeparator();
	long GetMaxDirDepth(void);
	void SetMaxDirDepth(long lDepth);

	void SetImportAsBackground(BOOL bImportAsBackground);
	void SetNavigateToPage(BOOL bNavigateToPage);
	void SetLimitPageWidth(BOOL bLimitPageWidth);
	void SetMaxPageWidth(float fMaxWidth);
	void SetIncludeSharedNotebooks(BOOL bIncludeSharedNotebooks);
	void SetFolderSeparator(LPCTSTR pszSeparator);

	// If this option is true, then we should rescan notebook each time user runs OneSnap.
	BOOL GetAutoSync(void);
	void SetAutoSync(BOOL bAutoSync);

private:
	CString LoadFromRegistry(LPCTSTR lpszValName, LPCTSTR lpszDefault);
	BOOL SaveToRegistry(LPCTSTR lpszValName, LPCTSTR lpszVal);
	void LoadRegSettings();
	void SaveRegSettings();
	// get the default path to the config file.   Default path is <NOTEBOOK PATH>\OneSnap.xml.
	CString GetDefaultFilepath(CString& strNotebookPath);
	// get the default path to "My Notebook".
	CString GetDefaultNotebookPath();

	void SetFilepath(LPCTSTR lpszFilepath);

	void AddCategoryEntry(LPCTSTR lpszPath, LPCTSTR lpszCategory, BOOL bIsNetworkPath, BOOL bFollowShortcuts, BOOL bFollowNetworkPaths);

	void AddCategories(LPCTSTR lpszPath, LPCTSTR lpszName, BOOL bFollowShortcuts, BOOL bFollowNetworkPaths, BOOL bPrependNameToSubs, LPCTSTR pszFolderSeparator, long nMaxDirDepth);
	void CreateConfigFile(LPCTSTR lpszFilepath, LPCTSTR lpszNotebookPath);

	CString m_strFilepath;
	CString m_strNotebookPath;

	void InitOptions();


};

