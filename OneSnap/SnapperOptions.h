#pragma once
#include "afxwin.h"

class CSnapperConfig;

// CSnapperOptions dialog

class CSnapperOptions : public CDialog
{
	DECLARE_DYNAMIC(CSnapperOptions)

public:
	CSnapperOptions(CWnd* pParent = NULL);   // standard constructor
	virtual ~CSnapperOptions();

	// modal processing
	INT_PTR DoModalEx(CSnapperConfig* pCfg);

	// Dialog Data
	enum { IDD = IDD_SNAPPEROPTIONS };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
private:
	afx_msg void OnBnClickedBrowseNotebook();
	afx_msg void OnBnClickedBrowseConfig();
	afx_msg void OnEnChangeConfigPath();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedUpdate();
	void InitDialog();
	void UpdateConfig(BOOL bFollowNetworkPaths = FALSE);
	virtual BOOL OnInitDialog();
	void InitToolTips();
	virtual BOOL PreTranslateMessage(MSG *pMsg);
	void WarnAutoSync();

	// set if we should (re-)synchronize the hotlists to the Notebook
	BOOL m_bSyncHotlists;
	// set if user wants main capture image imported as a page background.
	BOOL m_bBackground;
	// set if user wants OneNote to navigate to page after import
	BOOL m_bNavigateToPage;

	CString m_strNotebookPath;
	CString m_strConfigPath;
	CSnapperConfig*		m_pCfg;


	// when set update hotlists on Ok/Apply
	BOOL m_bUpdateHotlists;
	afx_msg void OnEnChangeNotebookPath();
	afx_msg void OnBnClickedApply();
	// set if we're supposed to limit the screen captures to the max width
	BOOL m_bLimitPageWidth;
	afx_msg void OnBnClickedIncshares();
	// when set hotlists should include shortcuts
	BOOL m_bIncludeShares;
	// max width when "limit page width" is set.
	CSliderCtrl m_sdrMaxWidth;
	afx_msg void OnNMCustomdrawWidth(NMHDR *pNMHDR, LRESULT *pResult);
	// max width when 'limit width' is set
	float m_fMaxWidth;
	afx_msg void OnBnClickedLimitWidth();
	// max width edit box
	CEdit m_edMaxWidth;
	afx_msg void OnEnChangeMaxwidth();
	// read-only text of the the minimum width
	CEdit m_edMaxWidthMin;
	// read-only control of the max width
	CEdit m_edMaxWidthMax;
	CToolTipCtrl m_ctrlToolTips;
	afx_msg void OnBnClickedAdvanced();
	afx_msg void OnBnClickedKeepSynced();
	// when true, rescan the notebook each time user runs OneSnap.
	BOOL m_bKeepSynced;
};
