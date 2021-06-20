// CRichEditCtrlEx.cpp: 实现文件
//

#include "pch.h"
#include "ExportTable.h"
#include "CRichEditCtrlEx.h"


// CRichEditCtrlEx

IMPLEMENT_DYNAMIC(CRichEditCtrlEx, CRichEditCtrl)

CRichEditCtrlEx::CRichEditCtrlEx()
{
}

CRichEditCtrlEx::~CRichEditCtrlEx()
{
}


BEGIN_MESSAGE_MAP(CRichEditCtrlEx, CRichEditCtrl)
    ON_NOTIFY_REFLECT(EN_LINK, &CRichEditCtrlEx::OnEnLink)
END_MESSAGE_MAP()



// CRichEditCtrlEx 消息处理程序




void CRichEditCtrlEx::OnEnLink(NMHDR* pNMHDR, LRESULT* pResult)
{
    ENLINK* pEnLink = reinterpret_cast<ENLINK*>(pNMHDR);
    // TODO: 在此添加控件通知处理程序代码
    if (pEnLink->msg == WM_LBUTTONDBLCLK)
    {
        SetSel(pEnLink->chrg);
        CString text = GetSelText();
        ShellExecute(NULL, _T("open"), text, _T(""), _T(""), SW_SHOWNORMAL);
    }
    *pResult = 0;
}
