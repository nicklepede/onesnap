// SnapperDialog.cpp : implementation file
//

#include "stdafx.h"
#include "OneSnap.h"
#include "SnapperDialog.h"
#include "SnapperOptions.h"
#include "SnapperConfig.h"

#undef GetFileTitle
// CSnapperDialog dialog

IMPLEMENT_DYNAMIC(CSnapperDialog, CDialog)

CSnapperDialog::CSnapperDialog(CWnd* pParent /*=NULL*/)
	: CDialog(CSnapperDialog::IDD, pParent)
	, m_strTitle(_T(""))
	, m_strComment(_T(""))
	, m_strCategory(_T(""))
	, m_strSection(_T(""))
{
	m_pCfg = NULL;
}

CSnapperDialog::~CSnapperDialog()
{
}

INT_PTR CSnapperDialog::DoModalEx(UserSettings* pSettings, CSnapperConfig* pCfg)
{
	m_pCfg = pCfg;
	SetUserSettings(pSettings);

	INT_PTR ret = DoModal();

	if (IDOK == ret)
	{
		GetUserSettings(pSettings);
		// this save is just to make sure we save if the use has added a new section by typing it in to the section list.  
		// this clearly isn't the right place for this...
		m_pCfg->Save();
	}

	m_pCfg = NULL;
	return ret;
}

BOOL CSnapperDialog::OnInitDialog()
{
	CDialog::OnInitDialog();

	InitDialog();
	return TRUE;
}

void CSnapperDialog::InitDialog()
{
	CStringArray rgHotlists;
	m_pCfg->GetHotlists(rgHotlists);
	PopulateCb(m_cbDirectory, rgHotlists);

	CString strCategory = m_pCfg->GetCurrentHotlist();
	m_cbDirectory.SelectString(-1, strCategory);

	CStringArray rgTargets;
	m_pCfg->GetTargets(rgTargets, strCategory);
	PopulateCb(m_cbSection, rgTargets);
	m_cbSection.SelectString(-1, m_pCfg->GetCurrentTarget(strCategory));
}

void CSnapperDialog::PopulateCb(CComboBox& cb, CStringArray& rgStrings)
{
	cb.ResetContent();
	for (int i = 0; i < rgStrings.GetSize(); i++)
		cb.AddString(rgStrings[i]);
}

void CSnapperDialog::GetUserSettings(UserSettings* psSettings)
{
	psSettings->strTitle = m_strTitle;
	psSettings->strComment = m_strComment;
	psSettings->strSectionFilepath = m_pCfg->GetPath(m_strCategory, m_strSection);
}
void CSnapperDialog::SetUserSettings(UserSettings* psSettings)
{
	m_strTitle = psSettings->strTitle;
	m_strComment = psSettings->strComment;
}
void CSnapperDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, ICB_DIRECTORY, m_cbDirectory);
	DDX_Control(pDX, ICB_SECTION, m_cbSection);
	DDX_Text(pDX, IED_TITLE, m_strTitle);
	DDX_Text(pDX, IED_COMMENT, m_strComment);
	DDX_CBString(pDX, ICB_DIRECTORY, m_strCategory);
	DDX_CBString(pDX, ICB_SECTION, m_strSection);
}


BEGIN_MESSAGE_MAP(CSnapperDialog, CDialog)
	ON_CBN_SELCHANGE(ICB_DIRECTORY, &CSnapperDialog::OnCbnSelchangeDirectory)
	ON_CBN_SELCHANGE(ICB_SECTION, &CSnapperDialog::OnCbnSelchangeSection)
	ON_BN_CLICKED(IBT_Configure, &CSnapperDialog::OnBnClickedConfigure)
	ON_BN_CLICKED(IDOK, &CSnapperDialog::OnBnClickedOk)
	ON_BN_CLICKED(IDBT_SCAN, &CSnapperDialog::OnBnClickedScan)
END_MESSAGE_MAP()

// configure the targets combo box, and save the default category back to the config file.
void CSnapperDialog::SelectCategory(CString& strCategory)
{

	CStringArray rgTargets;
	m_pCfg->GetTargets(rgTargets, strCategory);
	PopulateCb(m_cbSection, rgTargets);

	m_cbSection.SelectString(-1, m_pCfg->GetCurrentTarget(strCategory));

}

// CSnapperDialog message handlers
void CSnapperDialog::OnCbnSelchangeDirectory()
{
	UpdateData();
	SelectCategory(m_strCategory);
}

void CSnapperDialog::OnCbnSelchangeSection()
{
	UpdateData();
}

void CSnapperDialog::OnBnClickedConfigure()
{
	// save the current config.  This means we can throw away the options if
	// the user cancels, and we won't lose .  Not perfect, but oh well...
/*	UpdateData();
	CString strCategory = m_strCategory;
	CString strSection = m_strSection;
*/
	SaveConfig();
	CSnapperOptions dlgOptions;
	INT_PTR ret = dlgOptions.DoModalEx(m_pCfg);
	InitDialog();
	/*
	if (IDOK == ret)
	{
		try
		{
			SelectCategory(strCategory);
			if (0 == m_cbSection.SelectString(-1, strSection))
				throw "target no longer in hotlist";
		}
		catch(...)
		{
			// there was a problem w/ choosing the previously-selected hotlist/target,
			// so just use whatever the config has set.
			InitDialog();
		}
	}*/
}

void CSnapperDialog::SaveConfig()
{
	UpdateData();
	m_pCfg->SetCurrent(m_strCategory, m_strSection);
	m_pCfg->Save();
}

void CSnapperDialog::OnBnClickedOk()
{
	SaveConfig();
	OnOK();
}

void CSnapperDialog::OnBnClickedOk2()
{
	// TODO: Add your control notification handler code here
}

void CSnapperDialog::OnBnClickedScan()
{
	// TODO: Add your control notification handler code here
}
