#pragma once


// CSetupDialog 对话框

class CSetupDialog : public CDialog
{
	DECLARE_DYNAMIC(CSetupDialog)

public:
	CSetupDialog(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CSetupDialog();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_SETUP };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnBnClickedButtonSetup();
	afx_msg void OnBnClickedButtonChooseClientDir();
	afx_msg void OnBnClickedButtonChooseExcelDir();
	afx_msg void OnClose();

private:
	CString GetDirectory();

private:
	BOOL m_flag;
	CEdit m_excelDir;
	CEdit m_clientDir;
	CEdit m_hostId;
	CIPAddressCtrl m_serverIP;

	CString m_strExcelDir;
	CString m_strClientDir;
	BOOL m_bExcelDirChanged;
	CString m_strHostId;
	CString m_strServerIP;

};
