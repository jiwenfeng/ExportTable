// CListCtrlEx.cpp: 实现文件
//

#include "pch.h"
#include "ExportTable.h"
#include "CListCtrlEx.h"
#include <vector>


// CListCtrlEx

IMPLEMENT_DYNAMIC(CListCtrlEx, CListCtrl)

CListCtrlEx::CListCtrlEx() : m_ProgressColumn(0), m_oldTop(-1), m_oldLast(-1)
{
	m_chLastPressKey = ' ';
	m_nCount = -1;
	m_nLastPressTime = 0;
}

CListCtrlEx::~CListCtrlEx()
{
}


BEGIN_MESSAGE_MAP(CListCtrlEx, CListCtrl)
	ON_WM_PAINT()
	ON_WM_KEYDOWN()
	ON_NOTIFY_REFLECT(NM_KILLFOCUS, &CListCtrlEx::OnNMKillfocus)
	ON_WM_KILLFOCUS()
END_MESSAGE_MAP()

 //CListCtrlEx 消息处理程序
void CListCtrlEx::OnPaint()
{
	int Top = GetTopIndex();
	int Total = GetItemCount();
	int PerPage = GetCountPerPage();
	int LastItem = ((Top + PerPage) > Total) ? Total : Top + PerPage;
	int nCount = m_ProgressList.GetCount();
	CHeaderCtrl* pHeader = GetHeaderCtrl();

	if (m_oldTop != Top)
	{
		for (int i = 0; i < nCount; ++i)
		{
			CProgressBar* bar = m_ProgressList.GetAt(i);
			bar->HideWindow();
		}
	}
	for (int i = Top; i < LastItem; i++)
	{
		CProgressBar* bar = m_ProgressList.GetAt(i);
		CRect rt;
		GetItemRect(i, &rt, LVIR_LABEL);
		rt.top += 1;
		rt.bottom -= 1;
		CRect ColRt;
		pHeader->GetItemRect(m_ProgressColumn, &ColRt);
		rt.left += ColRt.left + 230;
		int Width = ColRt.Width();
		rt.right = rt.left + Width - 4 + 230;
		bar->ShowWindow(rt);
	}
	m_oldTop = Top;
	CListCtrl::OnPaint();
}

void CListCtrlEx::InitProgressColumn(int ColNum/*=0*/)
{
	m_ProgressColumn = ColNum;
	CHeaderCtrl* pHeader = GetHeaderCtrl();
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


void CListCtrlEx::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	if (nChar >= 'A' && nChar <= 'Z')
	{
		time_t now = time(NULL);
		if (now - m_nLastPressTime >= 30)
		{
			m_chLastPressKey = ' ';
		}
		m_nLastPressTime = time(NULL);
		if (m_chLastPressKey != nChar)
		{
			m_nCount = -1;
		}
		m_chLastPressKey = nChar;
		m_nCount++;
		for (int i = 0; i < GetItemCount(); ++i)
		{
			SetItemState(i, 0, LVIS_SELECTED);
		}
		std::vector<int> v;
		for (int i = 0; i < GetItemCount(); ++i)
		{
			CString strItem = GetItemText(i, 1);
			UINT ch = toupper(strItem.GetAt(0));
			if (ch == nChar)
			{
				v.push_back(i);
			}
		}
		if (v.size() != 0)
		{
			int row = m_nCount % v.size();
			SetItemState(row, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
			EnsureVisible(row, FALSE);
		}
	}
	CListCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CListCtrlEx::OnNMKillfocus(NMHDR* pNMHDR, LRESULT* pResult)
{
	// TODO: 在此添加控件通知处理程序代码
	m_nLastPressTime = 0;
	*pResult = 0;
}


void CListCtrlEx::OnKillFocus(CWnd* pNewWnd)
{
	m_nLastPressTime = 0;
	CListCtrl::OnKillFocus(pNewWnd);
	// TODO: 在此处添加消息处理程序代码
}