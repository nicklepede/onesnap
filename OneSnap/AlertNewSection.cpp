// AlertNewSection.cpp : implementation file
//

#include "stdafx.h"
#include "OneSnap.h"
#include "AlertNewSection.h"


// CAlertNewSection dialog

IMPLEMENT_DYNAMIC(CAlertNewSection, CDialog)

CAlertNewSection::CAlertNewSection(CWnd* pParent /*=NULL*/)
	: CDialog(CAlertNewSection::IDD, pParent)
	, m_bAlwaysSkip(FALSE)
{

}

CAlertNewSection::~CAlertNewSection()
{
}

void CAlertNewSection::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Check(pDX, IDC_ALWAYSSKIP, m_bAlwaysSkip);
}


BEGIN_MESSAGE_MAP(CAlertNewSection, CDialog)
END_MESSAGE_MAP()


// CAlertNewSection message handlers
