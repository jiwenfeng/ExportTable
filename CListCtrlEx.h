#pragma once


// CListCtrlEx
#define IDC_PROGRESS_LIST WM_USER

#include "CProgressBar.h"

class CListCtrlEx : public CListCtrl
{
	DECLARE_DYNAMIC(CListCtrlEx)

public:
	CListCtrlEx();
	virtual ~CListCtrlEx();

protected:
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnPaint();

	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);

	afx_msg void OnNMKillfocus(NMHDR* pNMHDR, LRESULT* pResult);

	afx_msg void OnKillFocus(CWnd* pNewWnd);

public:
	void InitProgressColumn(int ColNum);

	CProgressBar* GetProgressCtrl(int i);

private:
	int m_oldTop;
	int m_oldLast;
	int m_ProgressColumn;
	char m_chLastPressKey;
	char m_nCount;
	time_t m_nLastPressTime;
	CArray<CProgressBar*, CProgressBar*> m_ProgressList;
};


