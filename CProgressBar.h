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

	void DestroyWindow();

	void Invalidate(CExportTableDlg* window);

	CProgressCtrl* GetProgressCtl() { return m_pProgressCtrl; }

	void SetBarColor(COLORREF clrBar);
private:
	int m_show;
	int m_id;
	CWnd* m_parent;
	CProgressCtrl* m_pProgressCtrl;
};

