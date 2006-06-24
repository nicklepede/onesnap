#pragma once


// CAlertNewSection dialog

class CAlertNewSection : public CDialog
{
	DECLARE_DYNAMIC(CAlertNewSection)

public:
	CAlertNewSection(CWnd* pParent = NULL);   // standard constructor
	virtual ~CAlertNewSection();

// Dialog Data
	enum { IDD = IDD_NEWSECTION };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	// if user checks this then always skip this dialog
	BOOL m_bAlwaysSkip;
};
