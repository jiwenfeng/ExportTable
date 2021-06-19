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

	void InitProgressColumn(int ColNum);

	CProgressBar* GetProgressCtrl(int i);

private:
	CArray<CProgressBar*, CProgressBar *> m_ProgressList;
	// the column which should contain the progress bars
	int m_ProgressColumn;
};


