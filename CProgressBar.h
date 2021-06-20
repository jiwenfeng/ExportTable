#pragma once

class CExportTableDlg;
class CProgressBar
{
public:
	CProgressBar(int show, CWnd* parent, int id);

	~CProgressBar();

public:
	void ShowWindow(int show);

	void ShowWindow(const CRect& rect);

	void SetRange(short nLower, short nUpper);

	void SetPos(short pos);

	void OffsetPos(short pos);

	void HideWindow();

	void DestroyWindow();

	CProgressCtrl* GetProgressCtl() { return m_pProgressCtrl; }

	short GetRangeUpper();

	short GetPos();
private:
	int m_show;
	int m_id;
	short m_nLower;
	short m_nUpper;
	CWnd* m_parent;
	CProgressCtrl* m_pProgressCtrl;
};

