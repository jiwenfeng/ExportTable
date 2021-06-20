
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
#include "CSetupDialog.h"

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
	ON_COMMAND(ID_32772, &CExportTableDlg::On32772)
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

	HMENU menu = LoadMenu(NULL, MAKEINTRESOURCE(IDR_MENU_MAIN));
	::SetMenu(GetSafeHwnd(), menu);
	::DrawMenuBar(GetSafeHwnd());


	// TODO: 在此添加额外的初始化代码
	DWORD style = m_fileList.GetExStyle();
	style |= LVS_EX_CHECKBOXES;
	style |= LVS_EX_GRIDLINES;
	style |= LVS_EX_FULLROWSELECT;
	m_fileList.SetExtendedStyle(style);
	CRect rect;
	m_fileList.GetWindowRect(&rect);
	ScreenToClient(&rect);
	int width = rect.right - rect.left - 30;
	m_fileList.InsertColumn(0, _T(""), LVCFMT_CENTER, 50);
	m_fileList.InsertColumn(1, _T("文件名"), LVCFMT_LEFT,  200);
	m_fileList.InsertColumn(2, _T(""), LVCFMT_LEFT, width - 250);
	LoadConfig();
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExportTableDlg::LoadConfig()
{
	CString cfgFile = _T("./config.ini");
	wchar_t ed[MAX_PATH] = { 0 };
	wchar_t cd[MAX_PATH] = { 0 };
	wchar_t sd[MAX_PATH] = { 0 };
	wchar_t hd[MAX_PATH] = { 0 };
	GetPrivateProfileString(_T("Setup"), _T("Excel"), _T(""), ed, sizeof(ed), cfgFile);
	GetPrivateProfileString(_T("Setup"), _T("Client"), _T(""), cd, sizeof(cd), cfgFile);
	GetPrivateProfileString(_T("Setup"), _T("ServerIP"), _T(""), sd, sizeof(sd), cfgFile);
	GetPrivateProfileString(_T("Setup"), _T("HostID"), _T(""), hd, sizeof(hd), cfgFile);
	m_strExcelDir = ed;
	m_strClientDir = cd;
	m_strServerIP = sd;
	m_strHostID = hd;
	if (m_strClientDir.IsEmpty() || m_strExcelDir.IsEmpty() || m_strServerIP.IsEmpty() || m_strHostID.IsEmpty())
	{
		On32772();
		return;
	}
	CString fullPath = m_strExcelDir + _T("*.xls");
	LoadXLSFile(fullPath);
}

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
	m_fileList.DeleteAllItems();
	CFileFind ff;
	BOOL ret = ff.FindFile(path);
	int i = 0;
	while (ret)
	{
		ret = ff.FindNextFileW();
		if (ff.GetFileTitle() == "common_config")
		{
			continue;
		}
		m_fileList.InsertItem(i, _T(""));
		m_fileList.SetItemText(i, 1, ff.GetFileName());
		m_fileList.SetItemText(i, 2, _T(""));
		++i;
	}
	ff.Close();
	m_fileList.InitProgressColumn(i);
	if (m_fileList.GetItemCount() == 0) 
	{
		m_btnSelectAll.EnableWindow(FALSE);
		m_btnExport.EnableWindow(FALSE);
	}
	else
	{
		m_btnSelectAll.EnableWindow(TRUE);
		m_btnExport.EnableWindow(TRUE);
		m_fileList.SetFocus();
	}
}

int CExportTableDlg::CheckCallback(int id, const CString &file, const CString& sheet, const CString& err)
{
	m_richEdit.SetSel(-1, -1);
	long nStart = 0;
	long nEnd = 0;
	m_richEdit.GetSel(nStart, nEnd);
	m_richEdit.ReplaceSel(_T("处理[") + file + _T("]失败 ") + sheet + err + _T("\r\n"));
	CHARFORMAT2 cf;
	ZeroMemory(&cf, sizeof(cf));
	cf.cbSize = sizeof(CHARFORMAT2);
	cf.dwMask = CFM_COLOR | CFM_FACE | CFM_SIZE | CFM_LINK;
	cf.dwEffects |= CFE_LINK;

	m_richEdit.SetSel(nStart + 3, nStart + file.GetLength() + 3);
	m_richEdit.SetSelectionCharFormat(cf);
	m_richEdit.SendMessage(EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
	m_richEdit.SetEventMask(ENM_LINK);
	m_richEdit.SendMessage(WM_VSCROLL, SB_BOTTOM, 0);
	m_richEdit.SetSel(0, 0);
	if (m_status.find(id) == m_status.end())
	{
		m_status.insert(std::make_pair(id, 1));
	}
	return 0;
}

int CExportTableDlg::SetProgressBarCallback(int id, int nUpper)
{
	m_fileList.GetProgressCtrl(id)->SetRange(0, nUpper);
	m_fileList.SetItemText(id, 2, _T(""));
	return 0;
}

int CExportTableDlg::UpdateProgressBarCallback(int id, int nPos)
{
	m_fileList.GetProgressCtrl(id)->OffsetPos(1);

	if (m_fileList.GetProgressCtrl(id)->GetPos() == m_fileList.GetProgressCtrl(id)->GetRangeUpper())
	{
		m_fileList.GetProgressCtrl(id)->DestroyWindow();
		if (m_status.find(id) != m_status.end())
		{
			m_fileList.SetItemText(id, 2, _T("x"));
		}
		else
		{
			m_fileList.SetItemText(id, 2, _T("√"));
		}
	}
	else
	{
		m_fileList.SetItemText(id, 2, _T(""));
	}
	return 0;
}

void CExportTableDlg::CheckXLSFile(std::map<int, CString> exportList)
{
	m_status.clear();
	int nCount = m_fileList.GetItemCount();
	for (int i = 0; i < nCount; ++i)
	{
		m_fileList.SetItemText(i, 2, _T(""));
	}
	m_richEdit.SetWindowTextW(_T("开始检查数据表格式...\r\n"));
	for (std::map<int, CString>::iterator i = exportList.begin(); i != exportList.end(); ++i)
	{
		if (i->first > m_fileList.GetTopIndex())
		{
			int row = i->first + (m_fileList.GetCountPerPage() / 2);
			int realrow = row >= nCount ? nCount - 1 : row;
			m_fileList.EnsureVisible(realrow, FALSE);
		}
		else
		{
			m_fileList.EnsureVisible(i->first, FALSE);
		}


		CExcel excel(
			std::bind(&CExportTableDlg::CheckCallback, this, i->first, i->second, std::placeholders::_1, std::placeholders::_2)
		);
		excel.Check(
			i->second,
			std::bind(&CExportTableDlg::SetProgressBarCallback, this, i->first, std::placeholders::_1),
			std::bind(&CExportTableDlg::UpdateProgressBarCallback, this, i->first, std::placeholders::_1)
		);
		m_pgExport.OffsetPos(1);
	}

	m_richEdit.SetSel(-1, -1);
	if (!m_status.empty())
	{	
		m_richEdit.ReplaceSel(_T("数据表填写有误，导表终止\r\n"));
	}
	else
	{
		m_richEdit.ReplaceSel(_T("数据表格式检查完成，开始执行导表...\r\n"));
	}

	m_pgExport.SetPos(0);
	m_btnExport.EnableWindow(true);
	m_btnSelectAll.EnableWindow(true);
	m_flag = TRUE;
	OnBnClickedButtonSelectall();
	m_pgExport.ShowWindow(false);
}

void CExportTableDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	std::map<int, CString> exportList;
	for (int i = 0; i < m_fileList.GetItemCount(); ++i)
	{
		if (m_fileList.GetCheck(i))
		{
			CString file = m_strExcelDir + m_fileList.GetItemText(i, 1);
			exportList.insert(std::make_pair(i, file));
		}
	}
	if (exportList.empty())
	{
		MessageBox(_T("请选择要导出的文件"), _T("提示"), MB_ICONWARNING);
		return;
	}
	
	m_richEdit.SetWindowTextW(_T(""));
	m_btnExport.EnableWindow(false);
	m_btnSelectAll.EnableWindow(false);
	m_pgExport.ShowWindow(true);
	m_pgExport.SetRange(0, exportList.size());
	m_pgExport.SetPos(0);

	std::thread p(std::bind(&CExportTableDlg::CheckXLSFile, this, exportList));
	p.detach();
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
	CDialogEx::OnTimer(nIDEvent);
}

BOOL CExportTableDlg::DestroyWindow()
{
	// TODO: 在此添加专用代码和/或调用基类
	m_flag = 0;
	return CDialogEx::DestroyWindow();
}


void CExportTableDlg::On32772()
{
	// TODO: 在此添加命令处理程序代码
	CSetupDialog dialog;
	int ret = dialog.DoModal();
	if (ret == 1)
	{
		LoadConfig();
	}
}
