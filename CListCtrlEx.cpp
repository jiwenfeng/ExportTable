// CListCtrlEx.cpp: 实现文件
//

#include "pch.h"
#include "ExportTable.h"
#include "CListCtrlEx.h"


// CListCtrlEx

IMPLEMENT_DYNAMIC(CListCtrlEx, CListCtrl)

CListCtrlEx::CListCtrlEx() : m_ProgressColumn(0)
{

}

CListCtrlEx::~CListCtrlEx()
{
}


BEGIN_MESSAGE_MAP(CListCtrlEx, CListCtrl)
	ON_WM_PAINT()
END_MESSAGE_MAP()



// CListCtrlEx 消息处理程序
void CListCtrlEx::OnPaint()
{
	int Top = GetTopIndex();
	int Total = GetItemCount();
	int PerPage = GetCountPerPage();
	int LastItem = ((Top + PerPage) > Total) ? Total : Top + PerPage;
	for (int i = 0; i < m_ProgressList.GetCount(); ++i)
	{
		CProgressBar* bar = m_ProgressList.GetAt(i);
		bar->DestroyWindow();
	}
	CHeaderCtrl* pHeader = GetHeaderCtrl();
	for (int i = Top; i < LastItem; i++)
	{
		CProgressBar* bar = m_ProgressList.GetAt(i);
		CRect rt;
		GetItemRect(i, &rt, LVIR_LABEL);
		rt.top += 1;
		rt.bottom -= 1;
		CRect ColRt;
		pHeader->GetItemRect(m_ProgressColumn, &ColRt);
		rt.left += ColRt.left + 200;
		int Width = ColRt.Width();
		rt.right = rt.left + Width - 4 + 200;
		bar->ShowWindow(rt);
	}
	CListCtrl::OnPaint();
}


void CListCtrlEx::InitProgressColumn(int ColNum/*=0*/)
{
	m_ProgressColumn = ColNum;
	for (int i = 0; i < m_ProgressColumn; i++)
	{
		CProgressBar* pControl = new CProgressBar(0, this, IDC_PROGRESS_LIST + i);
		m_ProgressList.Add(pControl);
	}
}

CProgressBar* CListCtrlEx::GetProgressCtrl(int i)
{
	if (i > m_ProgressList.GetSize())
	{
		return NULL;
	}
	return m_ProgressList[i];
}
