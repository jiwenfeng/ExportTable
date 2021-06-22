
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
#include "CHelpDialog.h"
#include "curl/curl.h"
#include "json/json.h"

//#pragma comment(lib, "Shlwapi.lib")


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
	DDX_Control(pDX, IDC_EDIT_SERVER, m_serverEdit);
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
	ON_COMMAND(ID_32775, &CExportTableDlg::On32775)
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

	curl_global_init(CURL_GLOBAL_WIN32);

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

	int helped = GetPrivateProfileIntW(_T("Config"), _T("ShowHelp"), 0, cfgFile);
	if (!helped)
	{
		CHelpDialog dialog(this);
		dialog.DoModal();
	}

	wchar_t ed[MAX_PATH] = { 0 };
	wchar_t cd[MAX_PATH] = { 0 };
	wchar_t sd[MAX_PATH] = { 0 };
	wchar_t hd[MAX_PATH] = { 0 };
	GetPrivateProfileStringW(_T("Setup"), _T("Excel"), _T(""), ed, sizeof(ed), cfgFile);
	GetPrivateProfileStringW(_T("Setup"), _T("Client"), _T(""), cd, sizeof(cd), cfgFile);
	GetPrivateProfileStringW(_T("Setup"), _T("ServerIP"), _T(""), sd, sizeof(sd), cfgFile);
	GetPrivateProfileStringW(_T("Setup"), _T("HostID"), _T(""), hd, sizeof(hd), cfgFile);
	m_strExcelDir = ed;
	m_strClientDir = cd;
	m_strServerIP = sd;
	m_strHostID = hd;
	if (m_strClientDir.IsEmpty() || m_strExcelDir.IsEmpty() || m_strServerIP.IsEmpty() || m_strHostID.IsEmpty())
	{
		On32772();
		return;
	}
	LoadServer();
	LoadXLSFile();
}

void CExportTableDlg::LoadServer()
{
	CString cfgFile = _T("./config.ini");
	wchar_t sd[MAX_PATH] = { 0 };
	wchar_t hd[MAX_PATH] = { 0 };
	GetPrivateProfileStringW(_T("Setup"), _T("ServerIP"), _T(""), sd, sizeof(sd), cfgFile);
	GetPrivateProfileStringW(_T("Setup"), _T("HostID"), _T(""), hd, sizeof(hd), cfgFile);
	m_strServerIP = sd;
	m_strHostID = hd;
	m_serverEdit.SetWindowTextW(m_strServerIP + _T(":") + m_strHostID);
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

void CExportTableDlg::LoadXLSFile()
{
	CString cfgFile = _T("./config.ini");
	wchar_t ed[MAX_PATH] = { 0 };
	GetPrivateProfileStringW(_T("Setup"), _T("Excel"), _T(""), ed, sizeof(ed), cfgFile);
	m_strExcelDir = ed;

	m_fileList.DeleteAllItems();
	CFileFind ff;
	BOOL ret = ff.FindFile(m_strExcelDir + _T("*.xls"));
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

void CExportTableDlg::HttpResponse(void* ptr, size_t size, size_t nmemb, void* stream)
{
	char* str = (char*)ptr;
	CString* pstr = (CString*)stream;

	int bufSize = MultiByteToWideChar(CP_ACP, 0, str, -1, NULL, 0);
	wchar_t* pwstr = new wchar_t[bufSize];
	MultiByteToWideChar(CP_UTF8, 0, str, -1, pwstr, bufSize);
	std::wcout.imbue(std::locale("chs"));
	
	pstr->Append(pwstr, bufSize);
	delete[] pwstr;
}

void CExportTableDlg::Output(const CString& data)
{
	m_richEdit.SetSel(-1, -1);
	m_richEdit.ReplaceSel(data);
	m_richEdit.ReplaceSel(_T("\r\n"));
	m_richEdit.SendMessage(WM_VSCROLL, SB_BOTTOM, 0);
}

std::string CExportTableDlg::ConvertCStringToUTF8(CString strValue)
{
	std::wstring wbuffer;
#ifdef _UNICODE
	wbuffer.assign(strValue.GetString(), strValue.GetLength());
#else
	/*
	 * 转换ANSI到UNICODE
	 * 获取转换后长度
	 */
	int length = ::MultiByteToWideChar(CP_ACP, MB_ERR_INVALID_CHARS, (LPCTSTR)strValue, -1, NULL, 0);
	wbuffer.resize(length);
	/* 转换 */
	MultiByteToWideChar(CP_ACP, 0, (LPCTSTR)strValue, -1, (LPWSTR)(wbuffer.data()), wbuffer.length());
#endif

	/* 获取转换后长度 */
	int length = WideCharToMultiByte(CP_UTF8, 0, wbuffer.data(), wbuffer.size(), NULL, 0, NULL, NULL);
	/* 获取转换后内容 */
	std::string buffer;
	buffer.resize(length);

	WideCharToMultiByte(CP_UTF8, 0, strValue, -1, (LPSTR)(buffer.data()), length, NULL, NULL);
	return(buffer);
}

void CExportTableDlg::CheckXLSFile(std::map<int, CString> exportList)
{
	DWORD start = GetTickCount64();
	int nCount = m_fileList.GetItemCount();
	CExcel cmd(nullptr);
	const Table &table = cmd.Read(m_strExcelDir + _T("common_config.xls"), 1);
	std::map<CString, CString > psCmds;
	for (Table::const_iterator i = table.cbegin(); i != table.cend(); ++i)
	{
		CString cmd = i->at(0);
		CString name = i->at(1);
		psCmds.insert(std::make_pair(name, cmd));
	}
	//psCmds.insert(std::make_pair(_T("J_奖励表-英雄.xlsx"), _T("J_奖励表")));
	//psCmds.insert(std::make_pair(_T("J_奖励表.xlsx"), _T("J_奖励表")));
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

		if (i->second == _T("J_奖励表.xlsx") || i->second == _T("J_奖励表-英雄.xlsx"))
		{
			continue;
		}

		CExcel excel(
			std::bind(&CExportTableDlg::CheckCallback, this, i->first, m_strExcelDir + i->second, std::placeholders::_1, std::placeholders::_2)
		);
		excel.Check(
			m_strExcelDir + i->second,
			std::bind(&CExportTableDlg::SetProgressBarCallback, this, i->first, std::placeholders::_1),
			std::bind(&CExportTableDlg::UpdateProgressBarCallback, this, i->first, std::placeholders::_1)
		);
		m_pgExport.OffsetPos(1);
	}
	m_pgExport.SetPos(0);
	if (!m_status.empty())
	{	
		Output(_T("数据表填写有误，导表终止"));
	}
	else
	{
		Output(_T("数据表格式检查完成开始执行导表,请不要关闭程序..."));
		std::map<CString, int> psCache;
		int flag = TRUE;
		for (std::map<int, CString>::iterator i = exportList.begin(); i != exportList.end(); ++i)
		{
			CString strURL = m_strServerIP + _T(":") + m_strHostID;
			if (i->second == _T("J_奖励表.xlsx") || i->second == _T("J_奖励表-英雄.xlsx"))
			{
				int pos = i->second.ReverseFind('.');
				CString name = i->second.Left(pos);
				strURL += _T("/parse_reward?file=") + name;
			}
			else
			{
				std::map<CString, CString>::iterator itr = psCmds.find(i->second);
				if (itr == psCmds.end())
				{
					Output(_T("表格 [") + i->second + _T("] 没有找到ps指令,已跳过"));
					m_pgExport.OffsetPos(1);
					continue;
				}
				strURL += _T("/ps?file=") + itr->second;
				psCache[itr->second] = 1;
			}
			CURL* curl = curl_easy_init();
			std::string url = ConvertCStringToUTF8(strURL);
			CString response;
			curl_easy_setopt(curl, CURLOPT_URL, url.c_str());
			curl_easy_setopt(curl, CURLOPT_CONNECTTIMEOUT, 3);
			curl_easy_setopt(curl, CURLOPT_VERBOSE, 0);
			curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, CExportTableDlg::HttpResponse);
			curl_easy_setopt(curl, CURLOPT_WRITEDATA, (void*)&response);
			CURLcode res = curl_easy_perform(curl);
			curl_easy_cleanup(curl);
			if (res == CURLE_COULDNT_CONNECT)
			{
				Output(_T("致命错误！连接服务器" + m_strServerIP + _T(":") + m_strHostID + _T("失败，导表终止")));
				flag = FALSE;
				break;
			}
			std::string rsp = CExcel::CString2String(response);
			Json::Value root;
			Json::Reader reader;
			
			reader.parse(rsp.c_str(), root);
			if (root["IsSuccess"].asInt() == 1)
			{
				Output(_T("[") + i->second + _T("]导表成功!"));
			}
			else
			{
				Output(_T("[") + i->second + _T("]导表失败!"));
			}
			m_pgExport.OffsetPos(1);
		}
		if (flag)
		{
			ShellExecute(NULL, _T("open"), _T("copy.bat"), m_strExcelDir + _T(" ") + m_strClientDir, _T(""), SW_SHOWNORMAL);
		}
	}
	m_pgExport.SetPos(0);
	m_btnExport.EnableWindow(true);
	m_btnSelectAll.EnableWindow(true);
	m_flag = TRUE;
	OnBnClickedButtonSelectall();
	m_pgExport.ShowWindow(false);
	m_status.clear();
	for (int i = 0; i < nCount; ++i)
	{
		m_fileList.SetItemText(i, 2, _T(""));
	}

	DWORD finish = GetTickCount64();
	DWORD diff = finish - start;
	CString strCost = _T("导表结束，本次执行共用时:");
	if (diff > 60000)
	{
		CString strMin;
		strMin.Format(_T("%d分"), diff / 60000);
		strCost += strMin;
		diff %= 60000;
	}
	if (diff > 1000)
	{
		CString strMin;
		strMin.Format(_T("%d秒"), diff / 1000);
		strCost += strMin;
		diff %= 1000;
	}
	if (diff > 0)
	{
		CString strMin;
		strMin.Format(_T("%d毫秒"), diff);
		strCost += strMin;
	}
	Output(strCost);
}

void CExportTableDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	std::map<int, CString> exportList;
	for (int i = 0; i < m_fileList.GetItemCount(); ++i)
	{
		if (m_fileList.GetCheck(i))
		{
			exportList.insert(std::make_pair(i, m_fileList.GetItemText(i, 1)));
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
	CSetupDialog dialog(this);
	int ret = dialog.DoModal();
	if (ret & RELOAD_FILE_LIST)
	{
		LoadXLSFile();
	}
	if (ret & RELOAD_SERVER)
	{
		LoadServer();
	}
	CString cfgFile = _T("./config.ini");
	wchar_t cd[MAX_PATH] = { 0 };
	GetPrivateProfileStringW(_T("Setup"), _T("Client"), _T(""), cd, sizeof(cd), cfgFile);
	m_strClientDir = cd;
}

void CExportTableDlg::On32775()
{
	CHelpDialog dialog(this);
	dialog.DoModal();
	// TODO: 在此添加命令处理程序代码
}
