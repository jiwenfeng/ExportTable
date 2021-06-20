#pragma once


// CRichEditCtrlEx

class CRichEditCtrlEx : public CRichEditCtrl
{
	DECLARE_DYNAMIC(CRichEditCtrlEx)

public:
	CRichEditCtrlEx();
	virtual ~CRichEditCtrlEx();




protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnEnLink(NMHDR* pNMHDR, LRESULT* pResult);
};


