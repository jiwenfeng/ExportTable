
// ExportTableDlg.h: 头文件
//

#pragma once

#include <thread>
#include <map>
#include "CRichEditCtrlEx.h"
#include "CListCtrlEx.h"
#include <string>
#include <vector>

// CExportTableDlg 对话框
class CExportTableDlg : public CDialogEx
{
// 构造
public:
	CExportTableDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXPORTTABLE_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedCancel();
	afx_msg void OnLvnItemchangedListFile(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnBnClickedButtonSelectall();
	afx_msg void OnBnClickedOk();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	virtual BOOL DestroyWindow();
	afx_msg void On32772();
	afx_msg void On32775();

private:
	void LoadXLSFile();

	int CheckCallback(int id, const CString& file, const CString& sheet, const CString& err);

	int SetProgressBarCallback(int id, int nUpper);

	int UpdateProgressBarCallback(int id, int nPos);

	void CheckXLSFile(std::map<int, CString> mExportList);

	void DoExportTable(const std::map<int, CString>& mExportList, CFile* file, std::map<CString, std::vector<std::pair<CString, CString> > > &mExportFile);

	void LoadConfig();

	void LoadServer();

public:
	static void HttpResponse(void* ptr, size_t size, size_t nmemb, void* stream);

	void Output(const CString& data);

	std::string ConvertCStringToUTF8(CString strValue);

	BOOL HasRepeatFile(const std::map<CString, std::vector<std::pair<CString, CString> > >& mExportFiles, const CString& table, const std::vector<std::pair<CString, CString> >& vExportFiles);

private:
	CListCtrlEx m_fileList;
	CRichEditCtrlEx m_richEdit;
	CButton m_btnSelectAll;
	CButton m_btnExport;
	CProgressCtrl m_pgExport;
	CEdit m_serverEdit;

private:
	BOOL m_flag;
	CString m_strExcelDir;
	CString m_strClientDir;
	CString m_strServerIP;
	CString m_strHostID;
	std::map<int, int> m_status;
	std::map<CString, CString> m_cmds;
};
