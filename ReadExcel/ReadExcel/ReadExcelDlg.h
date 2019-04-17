
// ReadExcelDlg.h : 头文件
//

#pragma once
#include "afxwin.h"
#include "afxdtctl.h"


// CReadExcelDlg 对话框
class CReadExcelDlg : public CDialogEx
{
// 构造
public:
	CReadExcelDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_READEXCEL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnClickedReadexcel();
	afx_msg void OnClickedExportexcel();
	afx_msg void OnClickedButtonSaveIni();

	
	CString m_TimeMorningStart;
	CString m_TimeMorningEnd;
	CString m_TimeAfternoonStart;
	CString m_TimeAfternoonEnd;
	CString m_TimeEveningStart;
	CString m_TimeEveningEnd;
	CString m_WorkTimeMorningStart;
	CString m_WorkTimeMorningEnd;
	CString m_WorkTimeAfternoonStart;
	CString m_WorkTimeAfternoonEnd;
	CDateTimeCtrl m_startDateCtrl;
	CDateTimeCtrl m_endDateCtrl;

	vector<vector<CString>> vOriDatas;
	vector<vector<CString>> vResultDatas;
	vector<vector<CString>> vSummaryTables;
	int BelongToWhichTime(CString sTime);
	CString MorningStatus(vector<CString> vMorningTime, CString sWeekTime, int &iMorningTime);
	CString AfternoonStatus(vector<CString> vAfternoonTime, CString sWeekTime, int &iAfternoonTime);
	void EveningTimes(vector<CString> vEveningTime, int &iEveningTime);
	afx_msg void OnChangeEditMorningStart();
	afx_msg void OnChangeEditMorningEnd();
	afx_msg void OnChangeEditAfternoonStart();
	afx_msg void OnChangeEditAfternoonEnd();
	afx_msg void OnChangeEditEveningStart();
	afx_msg void OnChangeEditEveningEnd();
	afx_msg void OnChangeWorktimeEditMorningStart();
	afx_msg void OnChangeWorktimeEditMorningEnd();
	afx_msg void OnChangeWorktimeEditAfternoonStart();
	afx_msg void OnChangeWorktimeEditAfternoonEnd();
};
