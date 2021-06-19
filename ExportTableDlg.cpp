
// ExportTableDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "ExportTable.h"
#include "ExportTableDlg.h"
#include "afxdialogex.h"
#include <Shlwapi.h>
#include "Excel.h"
#include <memory>
#include <thread>
#include "Queue.h"

#pragma comment(lib, "Shlwapi.lib")


#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CExportTableDlg 对话框



CExportTableDlg::CExportTableDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_EXPORTTABLE_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_flag = FALSE;
}

void CExportTableDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_FILE, m_fileList);
	DDX_Control(pDX, IDC_RICHEDIT2_RESULT, m_richEdit);
	DDX_Control(pDX, IDC_BUTTON_SELECTALL, m_btnSelectAll);
	DDX_Control(pDX, IDOK, m_btnExport);
	DDX_Control(pDX, IDC_PROGRESS_EXPORT, m_pgExport);
}

BEGIN_MESSAGE_MAP(CExportTableDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDCANCEL, &CExportTableDlg::OnBnClickedCancel)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST_FILE, &CExportTableDlg::OnLvnItemchangedListFile)
	ON_BN_CLICKED(IDOK, &CExportTableDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON_SELECTALL, &CExportTableDlg::OnBnClickedButtonSelectall)
	ON_WM_TIMER()
END_MESSAGE_MAP()


// CExportTableDlg 消息处理程序

BOOL CExportTableDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	ShowWindow(SW_SHOWNORMAL);

	// TODO: 在此添加额外的初始化代码


	CString cfgFile = _T("./config.ini");

	BOOL flag = PathFileExists(cfgFile);
	if (!flag)
	{
		MessageBox(_T("没有找到config.init配置文件"), _T("错误"), MB_ICONERROR);
		return FALSE;
	}

	DWORD style = m_fileList.GetExStyle();
	style |= LVS_EX_CHECKBOXES;
	style |= LVS_EX_GRIDLINES;
	style |= LVS_EX_FULLROWSELECT;

	m_fileList.SetExtendedStyle(style);

	CRect rect;
	m_fileList.GetWindowRect(&rect);
	ScreenToClient(&rect);

	int width = rect.right - rect.left;


	m_fileList.InsertColumn(0, _T(""), LVCFMT_CENTER, 30);
	m_fileList.InsertColumn(1, _T("文件名"), LVCFMT_LEFT, width - 50);


	CString excelPath;

	int ret = GetPrivateProfileString(_T("dir"), _T("EXCEL_DIR"), _T(""), excelPath.GetBuffer(MAX_PATH), MAX_PATH, cfgFile);
	if (ret == 0)
	{
		MessageBox(_T("读取配置文件路径失败"), _T("错误"), MB_ICONERROR);
		return FALSE;
	}

	CString fullPath(excelPath.GetBuffer());

	m_excelPath = fullPath;
	
	fullPath.Append(L"\\*.*");
	
	LoadXLSFile(fullPath);

	std::thread p(std::bind(&CExportTableDlg::ShowResult, this));
	p.detach();

	std::thread p1(std::bind(&CExportTableDlg::ShowProgress, this));
	p1.detach();
	SetTimer(100, 100, NULL);
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExportTableDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExportTableDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CExportTableDlg::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnCancel();
}

void CExportTableDlg::OnLvnItemchangedListFile(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: 在此添加控件通知处理程序代码
	*pResult = 0;
}

void CExportTableDlg::LoadXLSFile(const CString& path)
{
	CFileFind ff;
	BOOL ret = ff.FindFile(path);

	int i = 0;
	while (ret)
	{
		ret = ff.FindNextFileW();
		if (ff.IsDirectory() || ff.IsDots() || ff.GetFileTitle() == "common_config")
		{
			continue;
		}

		CString fileName = ff.GetFileName();

		int pos = fileName.ReverseFind('.');
		if (-1 == pos)
		{
			continue;
		}
		CString ext = fileName.Right(fileName.GetLength() - pos - 1);
		if (ext != "xls")
		{
			continue;
		}
		m_fileList.InsertItem(i, _T(""));
		m_fileList.SetItemText(i++, 1, ff.GetFileName());
	}
	m_fileList.InitProgressColumn(i);
	ff.Close();
}

BOOL CExportTableDlg::ReadExcel(const CString& name, int id)
{
	std::shared_ptr<CExcel> pExcel(new CExcel);
	BOOL ret = pExcel->Check(name, id);
	if (!ret)
	{
		m_errList.push_back(id);
	}
	return TRUE;
}

void CExportTableDlg::CheckXLSFile()
{
	int nCount = m_fileList.GetItemCount();
	for (int i = 0; i < nCount; ++i)
	{
		if (m_fileList.GetCheck(i))
		{
			CString path = m_excelPath + "\\" + m_fileList.GetItemText(i, 1);
			ReadExcel(path, i);
			m_pgExport.SetPos(i);
		}
	}
	m_btnExport.EnableWindow(true);
	m_btnSelectAll.EnableWindow(true);
	m_flag = 1;
	OnBnClickedButtonSelectall();
	m_pgExport.ShowWindow(false);

	if (!m_errList.empty())
	{

	}
}

void CExportTableDlg::ShowProgress()
{
	while (TRUE)
	{
		ProgressCmd cmd = MQ<ProgressCmd>::getInstance().Pop();
		CProgressBar* bar = m_fileList.GetProgressCtrl(cmd.id);
		if (bar == NULL)
		{
			continue;
		}
		if (cmd.type == 0)
		{
			bar->SetRange(cmd.arg1, cmd.arg2);
			bar->ShowWindow(1);
			//bar->Invalidate(this);
		}
		else if (cmd.type == 1)
		{
			bar->SetPos(cmd.arg1);
			//bar->Invalidate(this);
		}
	}
}

void CExportTableDlg::ShowResult()
{
	while (TRUE)
	{
		CString data = Queue::getInstance().Pop();
		m_richEdit.SetSel(0, 0);
		m_richEdit.ReplaceSel(data + _T("\r\n"));
	}
}

void CExportTableDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	m_btnExport.EnableWindow(false);
	m_btnSelectAll.EnableWindow(false);
	//m_fileList.EnableWindow(false);
	m_pgExport.ShowWindow(true);
	m_pgExport.SetRange(0, m_fileList.GetItemCount());
	std::thread p(std::bind(&CExportTableDlg::CheckXLSFile, this));
	p.detach();
	//CDialogEx::OnOK();
}

void CExportTableDlg::OnBnClickedButtonSelectall()
{
	// TODO: 在此添加控件通知处理程序代码
	m_flag = !m_flag;
	if (m_flag)
	{
		m_btnSelectAll.SetWindowTextW(_T("取消全选"));
	}
	else
	{
		m_btnSelectAll.SetWindowTextW(_T("全选"));
	}
	int nCount = m_fileList.GetItemCount();
	for (int i = 0; i < nCount; ++i)
	{
		m_fileList.SetCheck(i, m_flag);
	}
}


void CExportTableDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	//m_fileList.OnPaint();
	CDialogEx::OnTimer(nIDEvent);
}


BOOL CExportTableDlg::DestroyWindow()
{
	// TODO: 在此添加专用代码和/或调用基类
	CDialogEx::DestroyWindow();
	exit(0);
}
