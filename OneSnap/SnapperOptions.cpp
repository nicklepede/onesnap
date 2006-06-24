// SnapperOptions.cpp : implementation file
//

#include "stdafx.h"
#include "OneSnap.h"
#include "SnapperOptions.h"
#include "SnapperConfig.h"

// CSnapperOptions dialog

IMPLEMENT_DYNAMIC(CSnapperOptions, CDialog)

CSnapperOptions::CSnapperOptions(CWnd* pParent /*=NULL*/)
	: CDialog(CSnapperOptions::IDD, pParent)
	, m_strNotebookPath(_T(""))
	, m_strConfigPath(_T(""))
	, m_bBackground(FALSE)
	, m_bNavigateToPage(FALSE)
	, m_bUpdateHotlists(FALSE)
	, m_bLimitPageWidth(FALSE)
	, m_bIncludeShares(FALSE)
	, m_fMaxWidth(0)
	, m_bKeepSynced(FALSE)
{

}

CSnapperOptions::~CSnapperOptions()
{
}

void CSnapperOptions::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDED_NOTEBOOK_PATH, m_strNotebookPath);
	DDX_Text(pDX, IDED_CONFIG_PATH, m_strConfigPath);
	DDX_Check(pDX, IDCHK_BACKGROUND, m_bBackground);
	DDX_Check(pDX, IDCHK_NAVIGATE, m_bNavigateToPage);
	DDX_Check(pDX, IDCHK_UPHOT, m_bUpdateHotlists);
	DDX_Check(pDX, IDCHK_LIMIT_WIDTH, m_bLimitPageWidth);
	DDX_Check(pDX, IDCHK_INCSHARES, m_bIncludeShares);
	DDX_Control(pDX, IDS_WIDTH, m_sdrMaxWidth);
	DDX_Text(pDX, IDED_MAXWIDTH, m_fMaxWidth);
	DDX_Control(pDX, IDED_MAXWIDTH, m_edMaxWidth);
	DDX_Control(pDX, IDED_6INCH, m_edMaxWidthMin);
	DDX_Control(pDX, IDED_36INCH, m_edMaxWidthMax);
	DDX_Check(pDX, IDC_AUTOSYNC, m_bKeepSynced);
}


BEGIN_MESSAGE_MAP(CSnapperOptions, CDialog)
	ON_BN_CLICKED(IDBT_BROWSE_NOTEBOOK, &CSnapperOptions::OnBnClickedBrowseNotebook)
	ON_BN_CLICKED(IDBT_BROWSE_CONFIG, &CSnapperOptions::OnBnClickedBrowseConfig)
	ON_EN_CHANGE(IDED_CONFIG_PATH, &CSnapperOptions::OnEnChangeConfigPath)
	ON_BN_CLICKED(IDOK, &CSnapperOptions::OnBnClickedOk)
	ON_BN_CLICKED(IDC_UPDATE, &CSnapperOptions::OnBnClickedUpdate)
	ON_EN_CHANGE(IDED_NOTEBOOK_PATH, &CSnapperOptions::OnEnChangeNotebookPath)
	ON_BN_CLICKED(IDAPPLY, &CSnapperOptions::OnBnClickedApply)
	ON_BN_CLICKED(IDCHK_INCSHARES, &CSnapperOptions::OnBnClickedIncshares)
	ON_NOTIFY(NM_CUSTOMDRAW, IDS_WIDTH, &CSnapperOptions::OnNMCustomdrawWidth)
	ON_BN_CLICKED(IDCHK_LIMIT_WIDTH, &CSnapperOptions::OnBnClickedLimitWidth)
	ON_EN_CHANGE(IDED_MAXWIDTH, &CSnapperOptions::OnEnChangeMaxwidth)
	ON_BN_CLICKED(IDC_ADVANCED, &CSnapperOptions::OnBnClickedAdvanced)
	ON_BN_CLICKED(IDC_AUTOSYNC, &CSnapperOptions::OnBnClickedKeepSynced)
END_MESSAGE_MAP()

INT_PTR CSnapperOptions::DoModalEx(CSnapperConfig* pCfg)
{
	ASSERT(pCfg != NULL);
	m_pCfg = pCfg;
	INT_PTR ret = DoModal();
	m_pCfg = NULL;

	return ret;
}

void CSnapperOptions::InitDialog()
{
	ASSERT(m_pCfg != NULL);
	m_strNotebookPath = m_pCfg->GetNotebookPath();
	m_strConfigPath	= m_pCfg->GetConfigFilePath();

	m_bBackground = m_pCfg->GetImportAsBackground();
	m_bNavigateToPage = m_pCfg->GetNavigateToPage();
	m_bLimitPageWidth = m_pCfg->GetLimitPageWidth();
	m_bIncludeShares = m_pCfg->GetIncludeSharedNotebooks();
	m_bKeepSynced = m_pCfg->GetAutoSync();
	m_fMaxWidth = m_pCfg->GetMaxPageWidth();
	m_edMaxWidthMin.SetWindowTextW(L"6\"");
	m_edMaxWidthMax.SetWindowTextW(L"36\"");
	
	m_sdrMaxWidth.SetRange(60, 360, FALSE);
	m_sdrMaxWidth.SetPos((int) (m_fMaxWidth * 10));
	m_sdrMaxWidth.EnableWindow(m_bLimitPageWidth);
	m_edMaxWidth.EnableWindow(m_bLimitPageWidth);
	m_edMaxWidthMin.EnableWindow(m_bLimitPageWidth);
	m_edMaxWidthMax.EnableWindow(m_bLimitPageWidth);

}

void CSnapperOptions::InitToolTips()
{
	m_ctrlToolTips.Create(this);	

	CWnd* pwndChild = GetWindow(GW_CHILD);
	CString strToolTip;
	while (pwndChild)
	{
		int idCtrl = pwndChild->GetDlgCtrlID();
		if (strToolTip.LoadString(idCtrl))
		{
			m_ctrlToolTips.AddTool(pwndChild, strToolTip);
		}
		pwndChild = pwndChild->GetWindow(GW_HWNDNEXT);
	}
	m_ctrlToolTips.Activate(TRUE);
}

BOOL CSnapperOptions::OnInitDialog()
{
	CDialog::OnInitDialog();
	InitToolTips();
	InitDialog();
	return TRUE;
}



// CSnapperOptions message handlers

void CSnapperOptions::OnBnClickedBrowseNotebook()
{
	BROWSEINFO sBi = { 0 };
	/*
    sBi.lpszTitle = TEXT("Select Notebook(s)");
	sBi.hwndOwner = m_hWndTop;
	sBi.iImage = 
	*/
    LPITEMIDLIST pIdl = SHBrowseForFolder ( &sBi );
    if ( pIdl != NULL )
    {
        // get the name of the folder
        TCHAR pszNotebookPath[MAX_PATH];
        if ( SHGetPathFromIDList ( pIdl, pszNotebookPath ) )
        {
			m_strNotebookPath = pszNotebookPath;
			m_bUpdateHotlists = TRUE;
			UpdateData(FALSE);
        }

        // free memory 
        IMalloc * piMalloc = NULL;
        if ( SUCCEEDED( SHGetMalloc ( &piMalloc )) )
        {
            piMalloc->Free ( pIdl );
            piMalloc->Release ( );
        }

    }
	
}

void CSnapperOptions::OnBnClickedBrowseConfig()
{
	CFileDialog	dlgConfigFile(TRUE, TEXT("xml"), m_strConfigPath/*TEXT("OneSnap.xml")*/, OFN_CREATEPROMPT | OFN_HIDEREADONLY, TEXT("XML Files (*.xml)|*.xml|All Files (*.*)|*.*||"));

	if (IDOK == dlgConfigFile.DoModal())
	{
		m_strConfigPath = dlgConfigFile.GetPathName();
		UpdateData(FALSE);
	}
//	strFilepath.ReleaseBuffer();
}

void CSnapperOptions::OnEnChangeConfigPath()
{

}

void CSnapperOptions::OnBnClickedOk()
{
	UpdateConfig();
	OnOK();
}

void CSnapperOptions::UpdateConfig(BOOL bFollowNetworkPaths /* = FALSE */)
{
	UpdateData(TRUE);
	m_pCfg->SetConfigFilePath(m_strConfigPath);
	m_pCfg->SetNotebookPath(m_strNotebookPath);
	m_pCfg->SetIncludeSharedNotebooks(m_bIncludeShares);
	m_pCfg->SetAutoSync(m_bKeepSynced);
	if (m_bUpdateHotlists)
	{
		BeginWaitCursor();
		m_pCfg->UpdateHotlists(bFollowNetworkPaths);
		m_bUpdateHotlists = FALSE;
		EndWaitCursor();
	}
	m_pCfg->SetImportAsBackground(m_bBackground);
	m_pCfg->SetNavigateToPage(m_bNavigateToPage);
	m_pCfg->SetLimitPageWidth(m_bLimitPageWidth);
	m_pCfg->SetMaxPageWidth(m_fMaxWidth);
	m_pCfg->Save();

}

void CSnapperOptions::OnBnClickedUpdate()
{
	UpdateData(TRUE);
	m_bUpdateHotlists = TRUE;
	UpdateData(FALSE);

	// if we've enabled shortcuts then warn the user it may take a while...
	if (m_bIncludeShares)
		MessageBox(_T("This may take a minute if any shortcuts point to inaccessible folders/sections."), _T("Starting Update"));

	// update the folder/section lists, following network paths.
	UpdateConfig(TRUE);
	// just in case we've loaded a new config file...
	InitDialog();
	UpdateData(FALSE);
	MessageBox(_T("Folder/section lists updated to match Notebook."), _T("Updated!"));
}

void CSnapperOptions::OnEnChangeNotebookPath()
{
	UpdateData(TRUE);
	m_bUpdateHotlists = TRUE;
	UpdateData(FALSE);
}

void CSnapperOptions::OnBnClickedApply()
{
	UpdateConfig();
	InitDialog();
	UpdateData(FALSE);
}

void CSnapperOptions::OnBnClickedIncshares()
{
	UpdateData(TRUE);
	m_bUpdateHotlists = TRUE;
	if (m_bKeepSynced && m_bIncludeShares)
		WarnAutoSync();
	UpdateData(FALSE);
}

void CSnapperOptions::WarnAutoSync()
{
	MessageBox(_T("When 'auto-syncing' OneSnap will not update shortcuts to remote folders/sections.  To update remote shortcuts you'll need to click 'Synchronize Now'."), _T("FYI"), MB_OK | MB_ICONWARNING );
}

void CSnapperOptions::OnNMCustomdrawWidth(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);

	m_fMaxWidth = ((float) m_sdrMaxWidth.GetPos()) / 10.0f;
	CString strWidth;
	UpdateData(FALSE);
	*pResult = 0;
}

void CSnapperOptions::OnBnClickedLimitWidth()
{
	UpdateData(TRUE);
	m_sdrMaxWidth.EnableWindow(m_bLimitPageWidth);
	m_edMaxWidth.EnableWindow(m_bLimitPageWidth);
	m_edMaxWidthMin.EnableWindow(m_bLimitPageWidth);
	m_edMaxWidthMax.EnableWindow(m_bLimitPageWidth);

}

void CSnapperOptions::OnEnChangeMaxwidth()
{
	UpdateData(TRUE);
	int iWidth = ((int) (m_fMaxWidth * 10));
	/*
	m_edMaxWidth.GetWindowText(strWidth);
	float fWidth = _tstof( strVal );
	int iWidth = ((int) (fWidth * 10);
	*/
	m_sdrMaxWidth.SetPos(iWidth);
}


// overloaded PreTranslateMessage to handle tooltips
BOOL CSnapperOptions::PreTranslateMessage(MSG *pMsg)
{
	m_ctrlToolTips.RelayEvent(pMsg);
	return CDialog::PreTranslateMessage(pMsg);
}


void CSnapperOptions::OnBnClickedAdvanced()
{
	
	// Open the OneSnap file that came w/ the plugin...
	// should be at <program files>\OneSnap\OneSnap.one

	CString strPath;
	// get the Program Files directory...
    if (!SUCCEEDED(SHGetFolderPath(NULL, CSIDL_PROGRAM_FILES, NULL, 0, strPath.GetBuffer(MAX_PATH))))
        if (0 == ExpandEnvironmentStrings(L"%ProgramFiles%\\", strPath.GetBuffer(MAX_PATH), MAX_PATH))
			strPath = L"C:\\Program Files\\";

	// add the filename...
	strPath += "\\OneSnap\\OneSnap.one";

	ShellExecute(NULL, L"open", L"C:\\Program Files\\OneSnap\\OneSnap.one", NULL, NULL, SW_SHOWNORMAL);
}


void CSnapperOptions::OnBnClickedKeepSynced()
{
	UpdateData(TRUE);
	if (m_bKeepSynced && m_bIncludeShares)
		WarnAutoSync();
	m_bUpdateHotlists = TRUE;
	UpdateData(FALSE);
}
