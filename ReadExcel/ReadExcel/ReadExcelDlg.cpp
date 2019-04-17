
// ReadExcelDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ReadExcel.h"
#include "ReadExcelDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CReadExcelDlg 对话框



CReadExcelDlg::CReadExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CReadExcelDlg::IDD, pParent)
	, m_TimeAfternoonEnd(_T("07:30"))
	, m_TimeMorningStart(_T("12:30"))
	, m_TimeMorningEnd(_T("12:30"))
	, m_TimeAfternoonStart(_T("18:00"))
	, m_TimeEveningStart(_T("18:00"))
	, m_TimeEveningEnd(_T("23:30"))
	, m_WorkTimeMorningStart(_T("08:20"))
	, m_WorkTimeMorningEnd(_T("11:30"))
	, m_WorkTimeAfternoonStart(_T("14:00"))
	, m_WorkTimeAfternoonEnd(_T("17:30"))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CReadExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_AFTERNOON_END, m_TimeAfternoonEnd);
	DDX_Text(pDX, IDC_EDIT_MORNING_START, m_TimeMorningStart);
	DDX_Text(pDX, IDC_EDIT_MORNING_END, m_TimeMorningEnd);
	DDX_Text(pDX, IDC_EDIT_AFTERNOON_START, m_TimeAfternoonStart);
	DDX_Text(pDX, IDC_EDIT_EVENING_START, m_TimeEveningStart);
	DDX_Text(pDX, IDC_EDIT_EVENING_END, m_TimeEveningEnd);
	DDX_Text(pDX, IDC_WORKTIME_EDIT_MORNING_START, m_WorkTimeMorningStart);
	DDX_Text(pDX, IDC_WORKTIME_EDIT_MORNING_END, m_WorkTimeMorningEnd);
	DDX_Text(pDX, IDC_WORKTIME_EDIT_AFTERNOON_START, m_WorkTimeAfternoonStart);
	DDX_Text(pDX, IDC_WORKTIME_EDIT_AFTERNOON_END, m_WorkTimeAfternoonEnd);	
	DDX_Control(pDX, IDC_STARTTIME, m_startDateCtrl);
	DDX_Control(pDX, IDC_ENDTIME, m_endDateCtrl);
}

BEGIN_MESSAGE_MAP(CReadExcelDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_ReadExcel, &CReadExcelDlg::OnClickedReadexcel)
	ON_BN_CLICKED(IDC_ExportExcel, &CReadExcelDlg::OnClickedExportexcel)
	ON_BN_CLICKED(IDC_Button_SaveIni, &CReadExcelDlg::OnClickedButtonSaveIni)
	ON_EN_CHANGE(IDC_EDIT_MORNING_START, &CReadExcelDlg::OnChangeEditMorningStart)
	ON_EN_CHANGE(IDC_EDIT_MORNING_END, &CReadExcelDlg::OnChangeEditMorningEnd)
	ON_EN_CHANGE(IDC_EDIT_AFTERNOON_START, &CReadExcelDlg::OnChangeEditAfternoonStart)
	ON_EN_CHANGE(IDC_EDIT_AFTERNOON_END, &CReadExcelDlg::OnChangeEditAfternoonEnd)
	ON_EN_CHANGE(IDC_EDIT_EVENING_START, &CReadExcelDlg::OnChangeEditEveningStart)
	ON_EN_CHANGE(IDC_EDIT_EVENING_END, &CReadExcelDlg::OnChangeEditEveningEnd)
	ON_EN_CHANGE(IDC_WORKTIME_EDIT_MORNING_START, &CReadExcelDlg::OnChangeWorktimeEditMorningStart)
	ON_EN_CHANGE(IDC_WORKTIME_EDIT_MORNING_END, &CReadExcelDlg::OnChangeWorktimeEditMorningEnd)
	ON_EN_CHANGE(IDC_WORKTIME_EDIT_AFTERNOON_START, &CReadExcelDlg::OnChangeWorktimeEditAfternoonStart)
	ON_EN_CHANGE(IDC_WORKTIME_EDIT_AFTERNOON_END, &CReadExcelDlg::OnChangeWorktimeEditAfternoonEnd)
END_MESSAGE_MAP()


// CReadExcelDlg 消息处理程序

BOOL CReadExcelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码
	CFileFind fileFinder;   //查找是否存在ini文件
	BOOL isFind = fileFinder.FindFile(_T(".\\AttendanceStatistic.ini"));
	if (!isFind)   // 如果文件不存在，则写配置文件
	{
		// 有效打卡时间配置
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningStart"), _T("07:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningEnd"), _T("12:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonStart"), _T("12:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonEnd"), _T("18:00"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningStart"), _T("18:00"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningEnd"), _T("23:30"), _T(".\\AttendanceStatistic.ini"));

		// 工作时间配置
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningStart"), _T("08:20"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningEnd"), _T("11:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonStart"), _T("14:00"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonEnd"), _T("17:30"), _T(".\\AttendanceStatistic.ini"));
		fileFinder.Close();
	}

	// 读配置文件，若对应值不存在，取指定的默认时间值
	// 读取ini配置文件  有效打卡时间配置
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeMorningStart"), _T("07:30"), m_TimeMorningStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeMorningEnd"), _T("12:30"), m_TimeMorningEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonStart"), _T("12:30"), m_TimeAfternoonStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonEnd"), _T("18:00"), m_TimeAfternoonEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeEveningStart"), _T("18:00"), m_TimeEveningStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeEveningEnd"), _T("23:30"), m_TimeEveningEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	// 工作时间配置
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningStart"), _T("08:20"), m_WorkTimeMorningStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningEnd"), _T("11:30"), m_WorkTimeMorningEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonStart"), _T("14:00"), m_WorkTimeAfternoonStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonEnd"), _T("17:30"), m_WorkTimeAfternoonEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));

	// 将配置显示在UI 有效打卡时间配置
	SetDlgItemText(IDC_EDIT_MORNING_START, m_TimeMorningStart);
	SetDlgItemText(IDC_EDIT_MORNING_END, m_TimeMorningEnd);
	SetDlgItemText(IDC_EDIT_AFTERNOON_START, m_TimeAfternoonStart);
	SetDlgItemText(IDC_EDIT_AFTERNOON_END, m_TimeAfternoonEnd);
	SetDlgItemText(IDC_EDIT_EVENING_START, m_TimeEveningStart);
	SetDlgItemText(IDC_EDIT_EVENING_END, m_TimeEveningEnd);
	// 工作时间配置
	SetDlgItemText(IDC_WORKTIME_EDIT_MORNING_START, m_WorkTimeMorningStart);
	SetDlgItemText(IDC_WORKTIME_EDIT_MORNING_END, m_WorkTimeMorningEnd);
	SetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_START, m_WorkTimeAfternoonStart);
	SetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_END, m_WorkTimeAfternoonEnd);


	//在这里必须ReleaseBuffer()，否则无法再后面进行字符串的连接  
	// 有效打卡时间配置
	m_TimeMorningStart.ReleaseBuffer();
	m_TimeMorningEnd.ReleaseBuffer();
	m_TimeAfternoonStart.ReleaseBuffer();
	m_TimeAfternoonEnd.ReleaseBuffer();
	m_TimeEveningStart.ReleaseBuffer();
	m_TimeEveningEnd.ReleaseBuffer();
	// 工作时间配置
	m_WorkTimeMorningStart.ReleaseBuffer();
	m_WorkTimeMorningEnd.ReleaseBuffer();
	m_WorkTimeAfternoonStart.ReleaseBuffer();
	m_WorkTimeAfternoonEnd.ReleaseBuffer();
	
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CReadExcelDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CReadExcelDlg::OnPaint()
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
HCURSOR CReadExcelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 获取ODBC中Excel驱动
CString GetExcelDriver()
{
	wchar_t szBuf[2001];
	wchar_t excl[] = L"Excel";
	WORD cbBufMax = 2000;
	WORD cbBufOut;
	wchar_t *pszBuf = szBuf;
	CString sDriver;
	// 获取已安装驱动的名称(函数在odbcinst.h里)
	if (!SQLGetInstalledDrivers(szBuf, cbBufMax, &cbBufOut))
		return L"";
	// 检索已安装的驱动是否有Excel...
	// AfxMessageBox(CString(pszBuf));
	do
	{
		if (wcsstr(pszBuf, excl) != 0)
		{
			// 找到驱动
			sDriver = CString(pszBuf);
			break;
		}
		wchar_t ze = { '\0' };
		pszBuf = wcschr(pszBuf, ze) + 1;
	} while (pszBuf[1] != '\0');

	return sDriver;
}

enum Times
{
	MORNING, AFTERNOON, EVENING, ERRORTIME
};

// 时间属于哪个时间段（MORNING, AFTERNOON, EVENING）
int CReadExcelDlg::BelongToWhichTime(CString sTime)
{
	if (sTime.IsEmpty())
	{
		return ERRORTIME;
	}

	COleDateTime sTempTime;
	sTempTime.ParseDateTime(sTime);

	COleDateTime morningStart;
	morningStart.ParseDateTime(m_TimeMorningStart);
	COleDateTime morningEnd;
	morningEnd.ParseDateTime(m_TimeMorningEnd);
	COleDateTime afternoonStart;
	afternoonStart.ParseDateTime(m_TimeAfternoonStart);
	COleDateTime afternoonEnd;
	afternoonEnd.ParseDateTime(m_TimeAfternoonEnd);
	COleDateTime eveningStart;
	eveningStart.ParseDateTime(m_TimeEveningStart);
	COleDateTime eveningEnd;
	eveningEnd.ParseDateTime(m_TimeEveningEnd);

	if (sTempTime >= morningStart && sTempTime <= morningEnd)
	{
		return MORNING;
	}
	else if (sTempTime > afternoonStart && sTempTime <= afternoonEnd)
	{
		return AFTERNOON;
	}
	else if (sTempTime > eveningStart && sTempTime <= eveningEnd)
	{
		return EVENING;
	}
	else
	{
		return ERRORTIME;
	}
}

// 判断上午情况
CString CReadExcelDlg::MorningStatus(vector<CString> vMorningTime, CString sWeekTime, int &iMorningTime)
{
	int iMorningTimeSize = vMorningTime.size();

	if (0 == iMorningTimeSize)
	{
		iMorningTime = 0;
		if (sWeekTime == "星期日")
		{
			return CString("正常");
		}
		else
		{
			return CString("未到");
		}
	}
	else if (1 == iMorningTimeSize)
	{
		iMorningTime = 0;
		if (sWeekTime == "星期日")
		{
			return CString("正常");
		}
		else
		{
			return CString("刷卡不足");
		}
	}
	else if (iMorningTimeSize == 2)
	{
		// 取出打卡时间
		COleDateTime sTempTimeStart;
		sTempTimeStart.ParseDateTime(vMorningTime[0]);
		COleDateTime sTempTimeEnd;
		sTempTimeEnd.ParseDateTime(vMorningTime[1]);
		// 星期日均计正常状态，直接计算时间
		if (sWeekTime == "星期日")
		{
			COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
			iMorningTime = (int)timeSpan.GetTotalMinutes();
			return CString("正常");
		}
		else  // 周一到周六计迟到早退
		{
			CString sMorningStart;
			CString sMorningEnd;
			if (sWeekTime == "星期六") // 周六打卡时间可稍晚
			{
				sMorningStart = "09:00";
				sMorningEnd = "11:30";
			}
			else
			{
				sMorningStart = m_WorkTimeMorningStart;
				sMorningEnd = m_WorkTimeMorningEnd;
			}
			COleDateTime workTimeMorningStart;
			workTimeMorningStart.ParseDateTime(sMorningStart);
			COleDateTime workTimeMorningEnd;
			workTimeMorningEnd.ParseDateTime(sMorningEnd);

			CString sTempStatus;
			if (sTempTimeStart <= workTimeMorningStart && sTempTimeEnd >= workTimeMorningEnd)
			{
				sTempStatus.Append(L"正常");
			}
			else
			{
				if (sTempTimeStart > workTimeMorningStart)
				{
					sTempStatus.Append(L"迟到");
				}

				if (sTempTimeEnd < workTimeMorningEnd)
				{
					sTempStatus.Append(L"早退");
				}
			}
			COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
			iMorningTime = (int)timeSpan.GetTotalMinutes();
			return sTempStatus;

		}
	}
	else
	{
		return CString("Error");
	}
	
}

// 判断下午情况
CString CReadExcelDlg::AfternoonStatus(vector<CString> vAfternoonTime, CString sWeekTime, int &iAfternoonTime)
{
	int iAfternoonTimeSize = vAfternoonTime.size();
	if (iAfternoonTimeSize == 0)
	{
		iAfternoonTime = 0;
		if (sWeekTime == "星期日")
		{
			return CString("正常");
		}
		else
		{
			return CString("未到");
		}
	}
	else if (iAfternoonTimeSize == 1)
	{
		iAfternoonTime = 0;
		if (sWeekTime == "星期日")
		{
			return CString("正常");
		}
		else
		{
			return CString("刷卡不足");
		}
	}
	else if (iAfternoonTimeSize == 2)
	{
		// 取出打卡时间
		COleDateTime sTempTimeStart;
		sTempTimeStart.ParseDateTime(vAfternoonTime[0]);
		COleDateTime sTempTimeEnd;
		sTempTimeEnd.ParseDateTime(vAfternoonTime[1]);
		// 星期日均计正常状态，直接计算时间
		if (sWeekTime == "星期日")
		{
			COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
			iAfternoonTime = (int)timeSpan.GetTotalMinutes();
			return CString("正常");
		}
		else // 周一到周六计迟到早退
		{
			CString sAfternoonStart;
			CString sAfternoonEnd;
			if (sWeekTime == "星期六") // 周六打卡时间可稍晚
			{
				sAfternoonStart = "14:30";
				sAfternoonEnd = "17:00";
			}
			else
			{
				sAfternoonStart = m_WorkTimeAfternoonStart;
				sAfternoonEnd = m_WorkTimeAfternoonEnd;
			}
			COleDateTime workTimeAfternoonStart;
			workTimeAfternoonStart.ParseDateTime(sAfternoonStart);
			COleDateTime workTimeAfternoonEnd;
			workTimeAfternoonEnd.ParseDateTime(sAfternoonEnd);
			CString sTempStatus;

			// 判断是否在中午打卡
			COleDateTime midAfternoonEnd;
			midAfternoonEnd.ParseDateTime(L"13:30");
			if (sTempTimeStart < midAfternoonEnd  && sTempTimeEnd < midAfternoonEnd) //下午两次打卡均在12:30--13:30
			{
				iAfternoonTime = 0;
				return CString("刷卡不足");
			}
			else if (sTempTimeStart < midAfternoonEnd  && sTempTimeEnd < workTimeAfternoonStart) // 第一次12:30--13:30 第二次13:30--14:00
			{
				iAfternoonTime = 0;
				return CString("刷卡不足");
			}
			else if (sTempTimeStart < midAfternoonEnd  && sTempTimeEnd > workTimeAfternoonStart) // 第一次12:30--13:30 第二次14:00后
			{
				sTempTimeStart = workTimeAfternoonStart;  // 如果在12:30-13:30之间打卡，按下午工作时间开始计
			}
			

			COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
			iAfternoonTime = (int)timeSpan.GetTotalMinutes();

			if (sTempTimeStart <= workTimeAfternoonStart && sTempTimeEnd >= workTimeAfternoonEnd)
			{
				sTempStatus.Append(L"正常");
			}
			else
			{
				if (sTempTimeStart > workTimeAfternoonStart)
				{
					sTempStatus.Append(L"迟到");
				}

				if (sTempTimeEnd < workTimeAfternoonEnd)
				{
					sTempStatus.Append(L"早退");
				}
			}
			return sTempStatus;
		}
	}
	else
	{
		return CString("Error");
	}	
}

// 计算晚上时间
void CReadExcelDlg::EveningTimes(vector<CString> vEveningTime, int &iEveningTime)
{
	int iMorningTimeSize = vEveningTime.size();
	if (iMorningTimeSize == 0 || iMorningTimeSize == 1)
	{
		iEveningTime = 0;
	}
	else if (iMorningTimeSize == 2)
	{
		COleDateTime sTempTimeStart;
		sTempTimeStart.ParseDateTime(vEveningTime[0]);
		COleDateTime sTempTimeEnd;
		sTempTimeEnd.ParseDateTime(vEveningTime[1]);
		COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
		iEveningTime = (int)timeSpan.GetTotalMinutes();
	}
}

void CReadExcelDlg::OnClickedReadexcel()
{
	// TODO:  在此添加控件通知处理程序代码
	// 设置过滤器   
	TCHAR szFilter[] = _T("表格文件(*.xls)|*.xls|所有文件(*.*)|*.*||");
	// 构造打开文件对话框   
	CFileDialog fileDlg(TRUE, _T("xls"), NULL, 0, szFilter, this);

	CString sPath;
	CString sFileName;
	// 显示打开文件对话框   
	if (IDOK == fileDlg.DoModal())
	{
		sPath = fileDlg.GetPathName();
		sFileName = fileDlg.GetFileName();
		// 如果点击了文件对话框上的“打开”按钮，则将选择的文件路径显示到编辑框里  
		SetDlgItemText(IDC_EDIT_FILEPATH, sPath);
	}
	else
	{
		return;
	}



	CDatabase DB;
	CString StrSQL;
	CString StrDsn;
	CString sDriver;

	// 检索是否安装有Excel驱动 "Microsoft Excel Driver (*.xls)" 
	sDriver = GetExcelDriver();
	if (sDriver.IsEmpty())
	{
		// 没有发现Excel驱动
		MessageBox(L"没有安装Excel驱动!");
		return;
	}

	//创建ODBC数据源连接字符串
	StrDsn.Format(L"ODBC;DRIVER={%s};DSN='';DBQ=%s", sDriver, sPath);
	TRY
	{
		//打开Excel文件
		DB.Open(NULL, false, false, StrDsn);
		CRecordset DBSet(&DB);

		//设置读取的查询语句
		CString sTableName = sFileName.Mid(0, sFileName.Find(L"."));  // 工作薄名与文件名相同
		StrSQL.Format(L"SELECT * FROM [%s$]", sTableName);
		
		//执行查询语句
		DBSet.Open(CRecordset::forwardOnly, StrSQL, CRecordset::readOnly);

		//获取查询结果
		//CString StrInfo = "";
		while (!DBSet.IsEOF())
		{
			//读取Excel内部数值
			vector<CString> vTempData;
			for (int i = 0; i< DBSet.GetODBCFieldCount(); i++)
			{
				CString Str;
				DBSet.GetFieldValue(i, Str);
				//StrInfo += Str + " ";
				vTempData.push_back(Str);
			}
			vOriDatas.push_back(vTempData);
			//StrInfo += "\n";
			DBSet.MoveNext();
		}
		//MessageBox(StrInfo);  //在信息框中显示
		MessageBox(L"数据读入成功!");
		GetDlgItem(IDC_ExportExcel)->EnableWindow(TRUE);
		DB.Close();
	}
	CATCH(CDBException, e)
	{
		MessageBox("数据库错误：" + e->m_strError);
	}
	END_CATCH;
}


void CReadExcelDlg::OnClickedExportexcel()
{
	// TODO:  在此添加控件通知处理程序代码
	// 设置过滤器   
	TCHAR szFilter[] = _T("表格文件(*.xls)|*.xls|所有文件(*.*)|*.*||");
	// 构造保存文件对话框   
	CFileDialog fileDlg(FALSE, _T("xls"), NULL, 0, szFilter, this);

	CString sPath;
	// 显示保存文件对话框   
	if (IDOK == fileDlg.DoModal())
	{
		sPath = fileDlg.GetPathName();
		if (PathFileExists(sPath))
		{
			CString sTempString;
			sTempString.Format(L"文件%s已存在，确定要覆盖该文件吗？", fileDlg.GetFileName());
			if (::MessageBox(NULL, sTempString, L"导出表格", MB_YESNO) == IDNO)
			{
				return;
			}
			DeleteFile(sPath);
		}
	}
	else
	{
		return;
	}
	
	/*  计算时间
	
	机构名称 人员编号	姓名	刷卡日期	星期	1	    2	   3	     4	     5	     6	   7
	cbib1	  00001     张三	2017-10-23 星期一 08:18   11:34   12:22    17:08    17:34  22:31

	*/
	int iOriDataSize = vOriDatas.size() - 1;  // 最后一条总记录数不计
	int iSumMinute = 0;
	for (int i = 0; i < iOriDataSize; i++)
	{
		vector<CString> vTempRawData = vOriDatas[i];
		int iTempRawDataSize = vTempRawData.size();
		vector<CString> vTempResultData;
		vector<CString> vTempMorningTime;
		vector<CString> vTempAfternoonTime;
		vector<CString> vTempEveningTime;
		for (int j = 1; j < iTempRawDataSize; j++)  //机构名称不需要 从1开始
		{
			CString sTempTime = vTempRawData[j];
			sTempTime.Trim();
			if (j >= 5)
			{
				//第5列开始 将时间分类
				switch (BelongToWhichTime(sTempTime))
				{
					case MORNING:
						if (vTempMorningTime.size() < 2)   // 超过2次打卡时，只记前两次
						{
							vTempMorningTime.push_back(sTempTime);
						}
						break;
					case AFTERNOON:
						if (vTempAfternoonTime.size() < 2)
						{
							vTempAfternoonTime.push_back(sTempTime);
						}
						break;
					case EVENING:
						if (vTempEveningTime.size() < 2)
						{
							vTempEveningTime.push_back(sTempTime);
						}
						break;
					default:
						break;
				} 
			}
			else
			{
				// 1-4列原样输出
				vTempResultData.push_back(sTempTime);
			}
		}

		// 上午、下午情况
		int iMorningTime = 0;
		int iAfternoonTime = 0;
		int iEveningTime = 0;
		CString sWeekTime = vTempResultData[3];
		vTempResultData.push_back((MorningStatus(vTempMorningTime, sWeekTime, iMorningTime)));
		vTempResultData.push_back((AfternoonStatus(vTempAfternoonTime, sWeekTime, iAfternoonTime)));
		EveningTimes(vTempEveningTime, iEveningTime);

		int iAllDayTime = iMorningTime + iAfternoonTime + iEveningTime;
		iSumMinute += iAllDayTime;

		CString sTempString;
		sTempString.Format(L"%d分钟", iMorningTime);
		vTempResultData.push_back(sTempString);
		sTempString.Format(L"%d分钟", iAfternoonTime);
		vTempResultData.push_back(sTempString);
		sTempString.Format(L"%d分钟", iEveningTime);
		vTempResultData.push_back(sTempString);
		sTempString.Format(L"%d分钟", iAllDayTime);
		vTempResultData.push_back(sTempString);

		// 将每行结果保存
		vResultDatas.push_back(vTempResultData);
		if (sWeekTime == "星期日") // 每位同学一周有七条记录，计算总时长
		{
			vector<CString> vSumTimes;
			sTempString.Format(L"总分钟:%d分钟  总小时:%d小时", iSumMinute, iSumMinute/60);
			vSumTimes.push_back(sTempString);
			vResultDatas.push_back(vSumTimes);

			vector<CString> vTempSummary;
			CString sPersonNum = vTempResultData[0];  // 人事编号
			CString sPersonName = vTempResultData[1]; // 姓名
			sTempString.Format(L"%d", iSumMinute/60);
			vTempSummary.push_back(sPersonNum);
			vTempSummary.push_back(sPersonName);
			vTempSummary.push_back(sTempString);
			vSummaryTables.push_back(vTempSummary);

			iSumMinute = 0;
		}
	}


	//创建Excel文件
	CDatabase DB;

	//Excel安装驱动
	CString StrDriver = "MICROSOFT EXCEL DRIVER (*.XLS)";

	//要建立的Execel文件
	CString StrExcelFile = sPath;
	CString StrSQL;
	StrSQL.Format(L"DRIVER={%s};DSN='';FIRSTROWHASNameS=1;READONLY=FALSE;CREATE_DB=%s;DBQ=%s", StrDriver, StrExcelFile, StrExcelFile);
	TRY
	{
		//创建Excel表格文件
		DB.OpenEx(StrSQL, CDatabase::noOdbcDialog);

		//创建表结构，字段名不能是Index
		StrSQL = "CREATE TABLE 学生考勤明细表(人事编号 TEXT, 姓名 TEXT, 刷卡日期 TEXT, 星期 TEXT, 上午情况 TEXT, 下午情况 TEXT, 上午时长 TEXT, 下午时长 TEXT, 晚上时长 TEXT, 一天总时长 TEXT)";
		DB.ExecuteSQL(StrSQL);

		//插入学生考勤明细表数据
		int iResultDatasSize = vResultDatas.size();
		for (int i = 0; i < iResultDatasSize; i++)
		{
			vector<CString> vTempRawData = vResultDatas[i];
			if (10 == vTempRawData.size())
			{
				StrSQL.Format(L"INSERT INTO 学生考勤明细表 (人事编号, 姓名, 刷卡日期, 星期, 上午情况, 下午情况, 上午时长, 下午时长, 晚上时长, 一天总时长) VALUES ('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')", vTempRawData[0], vTempRawData[1], vTempRawData[2], vTempRawData[3], vTempRawData[4], vTempRawData[5], vTempRawData[6], vTempRawData[7], vTempRawData[8], vTempRawData[9]);
				DB.ExecuteSQL(StrSQL);
			}
			else
			{
				StrSQL.Format(L"INSERT INTO 学生考勤明细表 (人事编号) VALUES ('%s')", vTempRawData[0]);
				DB.ExecuteSQL(StrSQL);
				if (i != iResultDatasSize - 1)
				{
					StrSQL.Format(L"INSERT INTO 学生考勤明细表 (人事编号) VALUES ('')");
					DB.ExecuteSQL(StrSQL);
					StrSQL.Format(L"INSERT INTO 学生考勤明细表 (人事编号, 姓名, 刷卡日期, 星期, 上午情况, 下午情况, 上午时长, 下午时长, 晚上时长, 一天总时长) VALUES ('人事编号', '姓名', '刷卡日期', '星期', '上午情况', '下午情况', '上午时长', '下午时长', '晚上时长', '一天总时长')");
					DB.ExecuteSQL(StrSQL);
				}
			}
		}

		// 插入CBIB周考勤总结表 
		StrSQL = "CREATE TABLE CBIB周考勤总结表(人事编号 TEXT, 姓名 TEXT, 总时长 TEXT, 备注 TEXT)";
		DB.ExecuteSQL(StrSQL);
		COleDateTime startDate;
		COleDateTime endDate;
		m_startDateCtrl.GetTime(startDate);
		m_endDateCtrl.GetTime(endDate);
		CString sStartDateTime;
		CString sEndDateTime;
		sStartDateTime.Format(L"%d/%d/%d", startDate.GetYear(), startDate.GetMonth(), startDate.GetDay());
		sEndDateTime.Format(L"%d/%d/%d", endDate.GetYear(), endDate.GetMonth(), endDate.GetDay());
		StrSQL.Format(L"INSERT INTO CBIB周考勤总结表 (人事编号) VALUES ('CBIB周考勤总结表 %s - %s')", sStartDateTime, sEndDateTime);
		DB.ExecuteSQL(StrSQL);

		int iSummaryTablesSize = vSummaryTables.size();
		for (int i = 0; i < iSummaryTablesSize; i++)
		{
			vector<CString> vTempRawData = vSummaryTables[i];
			StrSQL.Format(L"INSERT INTO CBIB周考勤总结表 (人事编号, 姓名, 总时长, 备注) VALUES ('%s', '%s', '%s', '')", vTempRawData[0], vTempRawData[1], vTempRawData[2]);
			DB.ExecuteSQL(StrSQL);
		}


		//关闭数据库
		DB.Close();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox("Excel表格创建错误：" + e->m_strError);
		return;
	}
	END_CATCH;

	MessageBox(L"导出成功！");
}


void CReadExcelDlg::OnClickedButtonSaveIni()
{
	// TODO:  在此添加控件通知处理程序代码
	CFileFind fileFinder;   //查找是否存在ini文件
	BOOL isFind = fileFinder.FindFile(_T(".\\AttendanceStatistic.ini"));
	if (isFind)
	{
		if (::MessageBox(NULL, L"确定要覆盖原有的配置吗？", L"保存配置", MB_YESNO) == IDNO)
		{
			return;
		}
	}

	// 有效打卡时间配置
	
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningStart"), m_TimeMorningStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningEnd"), m_TimeMorningEnd, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonStart"), m_TimeAfternoonStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonEnd"), m_TimeAfternoonEnd, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningStart"), m_TimeEveningStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningEnd"), m_TimeEveningEnd, _T(".\\AttendanceStatistic.ini"));

	// 工作时间配置
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningStart"), m_WorkTimeMorningStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningEnd"), m_WorkTimeMorningEnd, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonStart"), m_WorkTimeAfternoonStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonEnd"), m_WorkTimeAfternoonEnd, _T(".\\AttendanceStatistic.ini"));
	MessageBox(L"配置保存成功!");
	fileFinder.Close();

	// 全部重新初始化
	vOriDatas.clear();
	vResultDatas.clear();
	vSummaryTables.clear();
	GetDlgItem(IDC_ExportExcel)->EnableWindow(FALSE);
	SetDlgItemText(IDC_EDIT_FILEPATH, L"");
}

void CReadExcelDlg::OnChangeEditMorningStart()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_EDIT_MORNING_START, m_TimeMorningStart);
}


void CReadExcelDlg::OnChangeEditMorningEnd()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_EDIT_MORNING_END, m_TimeMorningEnd);
}


void CReadExcelDlg::OnChangeEditAfternoonStart()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_EDIT_AFTERNOON_START, m_TimeAfternoonStart);
}


void CReadExcelDlg::OnChangeEditAfternoonEnd()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_EDIT_AFTERNOON_END, m_TimeAfternoonEnd);
}


void CReadExcelDlg::OnChangeEditEveningStart()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_EDIT_EVENING_START, m_TimeEveningStart);
}


void CReadExcelDlg::OnChangeEditEveningEnd()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_EDIT_EVENING_END, m_TimeEveningEnd);
}


void CReadExcelDlg::OnChangeWorktimeEditMorningStart()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_WORKTIME_EDIT_MORNING_START, m_WorkTimeMorningStart);
}


void CReadExcelDlg::OnChangeWorktimeEditMorningEnd()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_WORKTIME_EDIT_MORNING_END, m_WorkTimeMorningEnd);
}


void CReadExcelDlg::OnChangeWorktimeEditAfternoonStart()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_START, m_WorkTimeAfternoonStart);
}


void CReadExcelDlg::OnChangeWorktimeEditAfternoonEnd()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	GetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_END, m_WorkTimeAfternoonEnd);
}
