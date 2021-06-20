// CHelpDialog.cpp: 实现文件
//

#include "pch.h"
#include "ExportTable.h"
#include "CHelpDialog.h"
#include "afxdialogex.h"


// CHelpDialog 对话框

IMPLEMENT_DYNAMIC(CHelpDialog, CDialog)

CHelpDialog::CHelpDialog(CWnd* pParent /*=nullptr*/)
	: CDialog(IDD_DIALOG_HELP, pParent)
{

}

CHelpDialog::~CHelpDialog()
{
}

void CHelpDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_RICHEDIT_DOC, m_docEdit);
}


BEGIN_MESSAGE_MAP(CHelpDialog, CDialog)
	ON_WM_SHOWWINDOW()
END_MESSAGE_MAP()


// CHelpDialog 消息处理程序


BOOL CHelpDialog::OnInitDialog()
{
	CDialog::OnInitDialog();

	// TODO:  在此添加额外的初始化
	HICON icon = AfxGetApp()->LoadIconW(IDR_MAINFRAME);
	SetIcon(icon, true);
	return TRUE;  // return TRUE unless you set the focus to a control
				  // 异常: OCX 属性页应返回 FALSE
}


void CHelpDialog::OnShowWindow(BOOL bShow, UINT nStatus)
{
	CString doc = _T("\
0、类型定义格式为name(type)，如果不是这个格式则不检查合法性\r\n\
1、数据类型支持int,float,string,str,mapp,array\r\n\
2、int 类型只能填写数字1-9，可以有负号但是不能是小数,如果要填入小数，请使用float类型\r\n\
3、string/str 类型可以使用\"包含，也可以不使用.如果使用\"\r\n\
4、map 类型必须要使用{}包含，如果其中包含字符串，必须要用\"包含\r\n\
5、array 类型可以使用[]包含也可以不适用，用,分割数组项。字符串可以不用\"包含\r\n\
6、第A列为索引列，不能重复，也不能为空\
");

	m_docEdit.SetSel(0, 0);
	m_docEdit.SetWindowTextW(doc);

	CHARFORMAT2 cf;
	ZeroMemory(&cf, sizeof(cf));
	cf.cbSize = sizeof(CHARFORMAT2);
	//cf.dwEffects |= CFE_BOLD;
	cf.yHeight = 300;                         //字体的大小（并非我们常见的字号概念）
	cf.dwMask = CFM_SIZE;
	m_docEdit.SetSel(0, doc.GetLength());
	m_docEdit.SetSelectionCharFormat(cf);
	m_docEdit.SendMessage(EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
	m_docEdit.SetSel(-1, -1);
	CDialog::OnShowWindow(bShow, nStatus);

	// TODO: 在此处添加消息处理程序代码
}
