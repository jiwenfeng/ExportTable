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
0、导表前请确保要导出的Excel已经提交到svn\r\n\
1、服务器IP填写要执行导表任务的服务器IP，端口是导表服务器的端口加5\r\n\
2、如果要修改设置，请点击菜单栏 文件->设置\r\n\
3、只支持后缀为xls文件的格式检查\r\n\
3、类型定义格式为name(type)，如果不是这个格式则不检查合法性，注意同一个sheet里面name不能重复且大小写敏感。举个栗子 id(int)\r\n\
4、数据类型支持int,float,string,str,map,array,macros,base64,float_base64\r\n\
5、int 类型只能填写整数，取值范围[-2147483648, 2147483647]\r\n\
6、要填小数请使用float类型，最多保留小数点后8位\r\n\
6、string/str 类型可以使用\"包含，也可以不使用。举个栗子 abc 或 \"abc\" 都是合法\r\n\
7、map 类型必须要使用{}包含，如果其中包含字符串，必须要用\"包含。举个栗子 {\"k1\":\"v1\", \"k2\":{\"kk1\":\"vv1\"}} 或 {123:243} 都是合法的。注意同一层级的key不能重复\r\n\
8、array 类型可以使用[]包含也可以不用。用,分割数组项。字符串可以不用\"包含。举个栗子 1,2,3 或者\"1\",\"2\",\"3\" 或 {},{},{} 或 [1,2,3,4]都是合法的\r\n\
9、第A列为索引列，不能重复，也不能为空\r\n\
10、不要包含空行\r\n\
11、要再次查看本文档请点击菜单栏 文件->帮助->关于\r\n\
12、奖励表暂时不进行数据检查。直接导表\r\n\
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
	m_docEdit.SetSel(0, 0);
	CDialog::OnShowWindow(bShow, nStatus);

	// TODO: 在此处添加消息处理程序代码
	CString cfgFile = _T("./config.ini");
	WritePrivateProfileString(_T("Config"), _T("ShowHelp"), _T("1"), cfgFile);
}
