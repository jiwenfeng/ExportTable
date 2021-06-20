// CSetupDialog.cpp: 实现文件
//

#include "pch.h"
#include "ExportTable.h"
#include "CSetupDialog.h"
#include "afxdialogex.h"


// CSetupDialog 对话框

IMPLEMENT_DYNAMIC(CSetupDialog, CDialog)



CSetupDialog::CSetupDialog(CWnd* pParent /*=nullptr*/)
	: CDialog(IDD_DIALOG_SETUP, pParent)
{
    m_flag = FALSE;
    m_status = 0;
}

CSetupDialog::~CSetupDialog()
{
}

void CSetupDialog::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_EDIT_EXCEL_DIR, m_excelDir);
    DDX_Control(pDX, IDC_EDIT_CLIENT_DIR, m_clientDir);
    DDX_Control(pDX, IDC_IPADDRESS_SERVER, m_serverIP);
    DDX_Control(pDX, IDC_EDIT_HOSTID, m_hostId);
}


BEGIN_MESSAGE_MAP(CSetupDialog, CDialog)
    ON_WM_SHOWWINDOW()
    ON_BN_CLICKED(IDC_BUTTON_SETUP, &CSetupDialog::OnBnClickedButtonSetup)
    ON_BN_CLICKED(IDC_BUTTON_CHOOSE_CLIENT_DIR, &CSetupDialog::OnBnClickedButtonChooseClientDir)
    ON_BN_CLICKED(IDC_BUTTON_CHOOSE_EXCEL_DIR, &CSetupDialog::OnBnClickedButtonChooseExcelDir)
    ON_WM_CLOSE()
END_MESSAGE_MAP()


// CSetupDialog 消息处理程序

CString CSetupDialog::GetDirectory()
{
    BROWSEINFO bi;
    wchar_t name[MAX_PATH];
    ZeroMemory(&bi, sizeof(BROWSEINFO));
    bi.hwndOwner = AfxGetMainWnd()->GetSafeHwnd();
    bi.pszDisplayName = name;
    bi.lpszTitle = _T("选择文件夹目录");
    bi.ulFlags = BIF_RETURNFSANCESTORS;
    LPITEMIDLIST idl = SHBrowseForFolder(&bi);
    if (idl == NULL)
        return _T("");
    CString strDirectoryPath;
    SHGetPathFromIDList(idl, strDirectoryPath.GetBuffer(MAX_PATH));
    strDirectoryPath.ReleaseBuffer();
    if (strDirectoryPath.IsEmpty())
        return _T("");
    if (strDirectoryPath.Right(1) != "\\")
        strDirectoryPath += "\\";

    return strDirectoryPath;
}


void CSetupDialog::OnShowWindow(BOOL bShow, UINT nStatus)
{
    ::SetWindowPos(this->GetSafeHwnd(), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    CDialog::OnShowWindow(bShow, nStatus);

    // TODO: 在此处添加消息处理程序代码
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
    m_strHostId = hd;
    m_excelDir.SetWindowTextW(m_strExcelDir);
    m_clientDir.SetWindowTextW(m_strClientDir);
    m_serverIP.SetWindowTextW(m_strServerIP);
    m_hostId.SetWindowTextW(m_strHostId);
    if (!m_strClientDir.IsEmpty() && !m_strExcelDir.IsEmpty() && !m_strServerIP.IsEmpty() && !m_strHostId.IsEmpty())
    {
        m_flag = TRUE;
    }
}

void CSetupDialog::OnBnClickedButtonSetup()
{
    // TODO: 在此添加控件通知处理程序代码
    CString strExcelDir, strClientDir, strServerIP, strHostId;
    m_excelDir.GetWindowTextW(strExcelDir);
    m_clientDir.GetWindowTextW(strClientDir);
    m_serverIP.GetWindowTextW(strServerIP);
    m_hostId.GetWindowTextW(strHostId);
    if (strExcelDir != m_strExcelDir && !strExcelDir.IsEmpty())
    {
        m_strExcelDir = strExcelDir;
        m_status |= RELOAD_FILE_LIST;
    }
    if (strClientDir != m_strClientDir && !strClientDir.IsEmpty())
    {
        m_strClientDir = strClientDir;
    }
    if (strServerIP != m_strServerIP && !strServerIP.IsEmpty())
    {
        m_strServerIP = strServerIP;
        m_status |= RELOAD_SERVER;
    }
    if (strHostId != m_strHostId && !strHostId.IsEmpty())
    {
        m_strHostId = strHostId;
        m_status |= RELOAD_SERVER;
    }
    if (m_strClientDir.IsEmpty() || m_strExcelDir.IsEmpty() || strServerIP.IsEmpty() || strHostId.IsEmpty())
    {
        MessageBox(_T("Excel文件夹、客户端文件夹、服务器IP和服务器ID不能为空"), _T("提示"), MB_ICONWARNING);
        return;
    }
    
    CString cfgFile = _T("./config.ini");
    WritePrivateProfileString(_T("Setup"), _T("Excel"), m_strExcelDir, cfgFile);
    WritePrivateProfileString(_T("Setup"), _T("Client"), m_strClientDir, cfgFile);
    WritePrivateProfileString(_T("Setup"), _T("ServerIP"), strServerIP, cfgFile);
    WritePrivateProfileString(_T("Setup"), _T("HostID"), strHostId, cfgFile);
    m_flag = TRUE;
    EndDialog(m_status);
}

void CSetupDialog::OnBnClickedButtonChooseClientDir()
{
    // TODO: 在此添加控件通知处理程序代码
    CString dir = GetDirectory();
    if (dir.IsEmpty())
    {
        return;
    }
    m_strClientDir = dir;
    m_clientDir.SetWindowTextW(dir);
}


void CSetupDialog::OnBnClickedButtonChooseExcelDir()
{
    // TODO: 在此添加控件通知处理程序代码
    CString dir = GetDirectory();
    if (dir.IsEmpty())
    {
        return;
    }
    if (dir != m_strExcelDir)
    {
        m_status = RELOAD_FILE_LIST;
    }
    m_strExcelDir = dir;
    m_excelDir.SetWindowTextW(dir);
}


void CSetupDialog::OnClose()
{
    // TODO: 在此添加消息处理程序代码和/或调用默认值
    if (!m_flag)
    {
        if (IDOK == MessageBox(_T("Excel文件夹、客户端文件夹、服务器IP和服务器ID不能为空, 是否退出"), _T("提示"), MB_OKCANCEL | MB_ICONWARNING))
        {
            exit(0);
        }
        return;
    }
    CDialog::OnClose();
}


BOOL CSetupDialog::OnInitDialog()
{
    CDialog::OnInitDialog();

    // TODO:  在此添加额外的初始化
    HICON icon = AfxGetApp()->LoadIconW(IDR_MAINFRAME);
    SetIcon(icon, true);
    return TRUE;  // return TRUE unless you set the focus to a control
                  // 异常: OCX 属性页应返回 FALSE
}
