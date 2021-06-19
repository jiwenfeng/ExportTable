
// ExportTableDlg.h: 头文件
//

#pragma once

#include <thread>
#include <vector>
#include "CListCtrlEx.h"

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


private:
	void LoadXLSFile(const CString &path);

	BOOL ReadExcel(const CString& name, int id);

	void CheckXLSFile();

	void ShowProgress();

	void ShowResult();

private:
	BOOL m_flag;
	CListCtrlEx m_fileList;
	CString m_excelPath;
	CRichEditCtrl m_richEdit;
	CButton m_btnSelectAll;
	CButton m_btnExport;
	CProgressCtrl m_pgExport;

private:
	std::vector<int> m_errList;
public:
	virtual BOOL DestroyWindow();
};
