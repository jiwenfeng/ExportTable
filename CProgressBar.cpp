#include "pch.h"
#include "CProgressBar.h"
#include "ExportTableDlg.h"

CProgressBar::CProgressBar(int show, CWnd *parent, int id)
{
	m_show = show;
	m_id = id;
	m_parent = parent;
	m_pProgressCtrl = new CProgressCtrl();
	CRect rect(0, 0, 0, 0);
	m_pProgressCtrl->Create(NULL, rect, parent, id);
}

CProgressBar::~CProgressBar()
{
	if (m_pProgressCtrl)
	{
		m_pProgressCtrl->DestroyWindow();
		delete m_pProgressCtrl;
		m_pProgressCtrl = NULL;
	}
}

void CProgressBar::ShowWindow(int show)
{
	m_show = show;
}

void CProgressBar::ShowWindow(const CRect &rect)
{
	m_pProgressCtrl->MoveWindow(rect);
	m_pProgressCtrl->ShowWindow(m_show);
}

void CProgressBar::SetRange(short nLower, short nUpper)
{
	m_show = 1;
	m_nLower = nLower;
	m_nUpper = nUpper;
	m_pProgressCtrl->SetRange(nLower, nUpper);
}

void CProgressBar::SetPos(short pos)
{
	m_pProgressCtrl->SetPos(pos);
}

void CProgressBar::OffsetPos(short pos)
{
	m_pProgressCtrl->OffsetPos(pos);
}

void CProgressBar::HideWindow()
{
	m_pProgressCtrl->ShowWindow(0);
}

void CProgressBar::DestroyWindow()
{
	if (m_show == 1)
	{
		m_show = 0;
		m_pProgressCtrl->ShowWindow(m_show);
		m_pProgressCtrl->SetPos(0);
	}
}

short CProgressBar::GetRangeUpper()
{
	return m_nUpper;
}

short CProgressBar::GetPos()
{
	return m_pProgressCtrl->GetPos();
}
