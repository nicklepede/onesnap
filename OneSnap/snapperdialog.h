#pragma once
#include "afxwin.h"
#include "SnapperXml.h"

class CSnapperConfig;
// SnapperDialog dialog

struct UserSettings
{
	CString strTitle;
	CString strComment;
	CString strSectionFilepath;
};

class CSnapperDialog : public CDialog
{
	DECLARE_DYNAMIC(CSnapperDialog)

public:
	CSnapperDialog(CWnd* pParent = NULL);   // standard constructor
	virtual ~CSnapperDialog();

	virtual BOOL OnInitDialog();
	
	// modal processing
	INT_PTR DoModalEx(UserSettings* pSettings, CSnapperConfig* pCfg);

// Dialog Data
	enum { IDD = IDD_SNAPPERDIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnCbnSelchangeDirectory();
	afx_msg void OnCbnSelchangeSection();
	CString m_strCategory;
	CString m_strSection;
	CString m_strFilepath;
	afx_msg void OnBnClickedConfigure();

private:
	void SelectCategory(CString& strCategory);
	void PopulateCb(CComboBox& cb, CStringArray& rgStrings);
	void SetUserSettings(UserSettings* psSettings);
	void GetUserSettings(UserSettings* psSettings);
	void InitDialog();
	void SaveConfig();

	CComboBox m_cbDirectory;
	CComboBox m_cbSection;
	CString m_strTitle;
	CString m_strComment;
	CString m_strUrl;
	CString m_strConfigFilepath;
	CString m_strSectionFilepath;
	CSnapperConfig* m_pCfg;

public:
	afx_msg void OnBnClickedOk();
public:
	afx_msg void OnBnClickedOk2();
public:
	afx_msg void OnBnClickedScan();
};
