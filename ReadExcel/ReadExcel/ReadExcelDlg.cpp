
// ReadExcelDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ReadExcel.h"
#include "ReadExcelDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CReadExcelDlg �Ի���



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


// CReadExcelDlg ��Ϣ�������

BOOL CReadExcelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������
	CFileFind fileFinder;   //�����Ƿ����ini�ļ�
	BOOL isFind = fileFinder.FindFile(_T(".\\AttendanceStatistic.ini"));
	if (!isFind)   // ����ļ������ڣ���д�����ļ�
	{
		// ��Ч��ʱ������
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningStart"), _T("07:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningEnd"), _T("12:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonStart"), _T("12:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonEnd"), _T("18:00"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningStart"), _T("18:00"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningEnd"), _T("23:30"), _T(".\\AttendanceStatistic.ini"));

		// ����ʱ������
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningStart"), _T("08:20"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningEnd"), _T("11:30"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonStart"), _T("14:00"), _T(".\\AttendanceStatistic.ini"));
		::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonEnd"), _T("17:30"), _T(".\\AttendanceStatistic.ini"));
		fileFinder.Close();
	}

	// �������ļ�������Ӧֵ�����ڣ�ȡָ����Ĭ��ʱ��ֵ
	// ��ȡini�����ļ�  ��Ч��ʱ������
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeMorningStart"), _T("07:30"), m_TimeMorningStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeMorningEnd"), _T("12:30"), m_TimeMorningEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonStart"), _T("12:30"), m_TimeAfternoonStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonEnd"), _T("18:00"), m_TimeAfternoonEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeEveningStart"), _T("18:00"), m_TimeEveningStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Effective Time"), _T("TimeEveningEnd"), _T("23:30"), m_TimeEveningEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	// ����ʱ������
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningStart"), _T("08:20"), m_WorkTimeMorningStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningEnd"), _T("11:30"), m_WorkTimeMorningEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonStart"), _T("14:00"), m_WorkTimeAfternoonStart.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));
	::GetPrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonEnd"), _T("17:30"), m_WorkTimeAfternoonEnd.GetBuffer(MAX_PATH), MAX_PATH, _T(".\\AttendanceStatistic.ini"));

	// ��������ʾ��UI ��Ч��ʱ������
	SetDlgItemText(IDC_EDIT_MORNING_START, m_TimeMorningStart);
	SetDlgItemText(IDC_EDIT_MORNING_END, m_TimeMorningEnd);
	SetDlgItemText(IDC_EDIT_AFTERNOON_START, m_TimeAfternoonStart);
	SetDlgItemText(IDC_EDIT_AFTERNOON_END, m_TimeAfternoonEnd);
	SetDlgItemText(IDC_EDIT_EVENING_START, m_TimeEveningStart);
	SetDlgItemText(IDC_EDIT_EVENING_END, m_TimeEveningEnd);
	// ����ʱ������
	SetDlgItemText(IDC_WORKTIME_EDIT_MORNING_START, m_WorkTimeMorningStart);
	SetDlgItemText(IDC_WORKTIME_EDIT_MORNING_END, m_WorkTimeMorningEnd);
	SetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_START, m_WorkTimeAfternoonStart);
	SetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_END, m_WorkTimeAfternoonEnd);


	//���������ReleaseBuffer()�������޷��ٺ�������ַ���������  
	// ��Ч��ʱ������
	m_TimeMorningStart.ReleaseBuffer();
	m_TimeMorningEnd.ReleaseBuffer();
	m_TimeAfternoonStart.ReleaseBuffer();
	m_TimeAfternoonEnd.ReleaseBuffer();
	m_TimeEveningStart.ReleaseBuffer();
	m_TimeEveningEnd.ReleaseBuffer();
	// ����ʱ������
	m_WorkTimeMorningStart.ReleaseBuffer();
	m_WorkTimeMorningEnd.ReleaseBuffer();
	m_WorkTimeAfternoonStart.ReleaseBuffer();
	m_WorkTimeAfternoonEnd.ReleaseBuffer();
	
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CReadExcelDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CReadExcelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// ��ȡODBC��Excel����
CString GetExcelDriver()
{
	wchar_t szBuf[2001];
	wchar_t excl[] = L"Excel";
	WORD cbBufMax = 2000;
	WORD cbBufOut;
	wchar_t *pszBuf = szBuf;
	CString sDriver;
	// ��ȡ�Ѱ�װ����������(������odbcinst.h��)
	if (!SQLGetInstalledDrivers(szBuf, cbBufMax, &cbBufOut))
		return L"";
	// �����Ѱ�װ�������Ƿ���Excel...
	// AfxMessageBox(CString(pszBuf));
	do
	{
		if (wcsstr(pszBuf, excl) != 0)
		{
			// �ҵ�����
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

// ʱ�������ĸ�ʱ��Σ�MORNING, AFTERNOON, EVENING��
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

// �ж��������
CString CReadExcelDlg::MorningStatus(vector<CString> vMorningTime, CString sWeekTime, int &iMorningTime)
{
	int iMorningTimeSize = vMorningTime.size();

	if (0 == iMorningTimeSize)
	{
		iMorningTime = 0;
		if (sWeekTime == "������")
		{
			return CString("����");
		}
		else
		{
			return CString("δ��");
		}
	}
	else if (1 == iMorningTimeSize)
	{
		iMorningTime = 0;
		if (sWeekTime == "������")
		{
			return CString("����");
		}
		else
		{
			return CString("ˢ������");
		}
	}
	else if (iMorningTimeSize == 2)
	{
		// ȡ����ʱ��
		COleDateTime sTempTimeStart;
		sTempTimeStart.ParseDateTime(vMorningTime[0]);
		COleDateTime sTempTimeEnd;
		sTempTimeEnd.ParseDateTime(vMorningTime[1]);
		// �����վ�������״̬��ֱ�Ӽ���ʱ��
		if (sWeekTime == "������")
		{
			COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
			iMorningTime = (int)timeSpan.GetTotalMinutes();
			return CString("����");
		}
		else  // ��һ�������Ƴٵ�����
		{
			CString sMorningStart;
			CString sMorningEnd;
			if (sWeekTime == "������") // ������ʱ�������
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
				sTempStatus.Append(L"����");
			}
			else
			{
				if (sTempTimeStart > workTimeMorningStart)
				{
					sTempStatus.Append(L"�ٵ�");
				}

				if (sTempTimeEnd < workTimeMorningEnd)
				{
					sTempStatus.Append(L"����");
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

// �ж��������
CString CReadExcelDlg::AfternoonStatus(vector<CString> vAfternoonTime, CString sWeekTime, int &iAfternoonTime)
{
	int iAfternoonTimeSize = vAfternoonTime.size();
	if (iAfternoonTimeSize == 0)
	{
		iAfternoonTime = 0;
		if (sWeekTime == "������")
		{
			return CString("����");
		}
		else
		{
			return CString("δ��");
		}
	}
	else if (iAfternoonTimeSize == 1)
	{
		iAfternoonTime = 0;
		if (sWeekTime == "������")
		{
			return CString("����");
		}
		else
		{
			return CString("ˢ������");
		}
	}
	else if (iAfternoonTimeSize == 2)
	{
		// ȡ����ʱ��
		COleDateTime sTempTimeStart;
		sTempTimeStart.ParseDateTime(vAfternoonTime[0]);
		COleDateTime sTempTimeEnd;
		sTempTimeEnd.ParseDateTime(vAfternoonTime[1]);
		// �����վ�������״̬��ֱ�Ӽ���ʱ��
		if (sWeekTime == "������")
		{
			COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
			iAfternoonTime = (int)timeSpan.GetTotalMinutes();
			return CString("����");
		}
		else // ��һ�������Ƴٵ�����
		{
			CString sAfternoonStart;
			CString sAfternoonEnd;
			if (sWeekTime == "������") // ������ʱ�������
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

			// �ж��Ƿ��������
			COleDateTime midAfternoonEnd;
			midAfternoonEnd.ParseDateTime(L"13:30");
			if (sTempTimeStart < midAfternoonEnd  && sTempTimeEnd < midAfternoonEnd) //�������δ򿨾���12:30--13:30
			{
				iAfternoonTime = 0;
				return CString("ˢ������");
			}
			else if (sTempTimeStart < midAfternoonEnd  && sTempTimeEnd < workTimeAfternoonStart) // ��һ��12:30--13:30 �ڶ���13:30--14:00
			{
				iAfternoonTime = 0;
				return CString("ˢ������");
			}
			else if (sTempTimeStart < midAfternoonEnd  && sTempTimeEnd > workTimeAfternoonStart) // ��һ��12:30--13:30 �ڶ���14:00��
			{
				sTempTimeStart = workTimeAfternoonStart;  // �����12:30-13:30֮��򿨣������繤��ʱ�俪ʼ��
			}
			

			COleDateTimeSpan timeSpan = sTempTimeEnd - sTempTimeStart;
			iAfternoonTime = (int)timeSpan.GetTotalMinutes();

			if (sTempTimeStart <= workTimeAfternoonStart && sTempTimeEnd >= workTimeAfternoonEnd)
			{
				sTempStatus.Append(L"����");
			}
			else
			{
				if (sTempTimeStart > workTimeAfternoonStart)
				{
					sTempStatus.Append(L"�ٵ�");
				}

				if (sTempTimeEnd < workTimeAfternoonEnd)
				{
					sTempStatus.Append(L"����");
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

// ��������ʱ��
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
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	// ���ù�����   
	TCHAR szFilter[] = _T("����ļ�(*.xls)|*.xls|�����ļ�(*.*)|*.*||");
	// ������ļ��Ի���   
	CFileDialog fileDlg(TRUE, _T("xls"), NULL, 0, szFilter, this);

	CString sPath;
	CString sFileName;
	// ��ʾ���ļ��Ի���   
	if (IDOK == fileDlg.DoModal())
	{
		sPath = fileDlg.GetPathName();
		sFileName = fileDlg.GetFileName();
		// ���������ļ��Ի����ϵġ��򿪡���ť����ѡ����ļ�·����ʾ���༭����  
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

	// �����Ƿ�װ��Excel���� "Microsoft Excel Driver (*.xls)" 
	sDriver = GetExcelDriver();
	if (sDriver.IsEmpty())
	{
		// û�з���Excel����
		MessageBox(L"û�а�װExcel����!");
		return;
	}

	//����ODBC����Դ�����ַ���
	StrDsn.Format(L"ODBC;DRIVER={%s};DSN='';DBQ=%s", sDriver, sPath);
	TRY
	{
		//��Excel�ļ�
		DB.Open(NULL, false, false, StrDsn);
		CRecordset DBSet(&DB);

		//���ö�ȡ�Ĳ�ѯ���
		CString sTableName = sFileName.Mid(0, sFileName.Find(L"."));  // �����������ļ�����ͬ
		StrSQL.Format(L"SELECT * FROM [%s$]", sTableName);
		
		//ִ�в�ѯ���
		DBSet.Open(CRecordset::forwardOnly, StrSQL, CRecordset::readOnly);

		//��ȡ��ѯ���
		//CString StrInfo = "";
		while (!DBSet.IsEOF())
		{
			//��ȡExcel�ڲ���ֵ
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
		//MessageBox(StrInfo);  //����Ϣ������ʾ
		MessageBox(L"���ݶ���ɹ�!");
		GetDlgItem(IDC_ExportExcel)->EnableWindow(TRUE);
		DB.Close();
	}
	CATCH(CDBException, e)
	{
		MessageBox("���ݿ����" + e->m_strError);
	}
	END_CATCH;
}


void CReadExcelDlg::OnClickedExportexcel()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	// ���ù�����   
	TCHAR szFilter[] = _T("����ļ�(*.xls)|*.xls|�����ļ�(*.*)|*.*||");
	// ���챣���ļ��Ի���   
	CFileDialog fileDlg(FALSE, _T("xls"), NULL, 0, szFilter, this);

	CString sPath;
	// ��ʾ�����ļ��Ի���   
	if (IDOK == fileDlg.DoModal())
	{
		sPath = fileDlg.GetPathName();
		if (PathFileExists(sPath))
		{
			CString sTempString;
			sTempString.Format(L"�ļ�%s�Ѵ��ڣ�ȷ��Ҫ���Ǹ��ļ���", fileDlg.GetFileName());
			if (::MessageBox(NULL, sTempString, L"�������", MB_YESNO) == IDNO)
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
	
	/*  ����ʱ��
	
	�������� ��Ա���	����	ˢ������	����	1	    2	   3	     4	     5	     6	   7
	cbib1	  00001     ����	2017-10-23 ����һ 08:18   11:34   12:22    17:08    17:34  22:31

	*/
	int iOriDataSize = vOriDatas.size() - 1;  // ���һ���ܼ�¼������
	int iSumMinute = 0;
	for (int i = 0; i < iOriDataSize; i++)
	{
		vector<CString> vTempRawData = vOriDatas[i];
		int iTempRawDataSize = vTempRawData.size();
		vector<CString> vTempResultData;
		vector<CString> vTempMorningTime;
		vector<CString> vTempAfternoonTime;
		vector<CString> vTempEveningTime;
		for (int j = 1; j < iTempRawDataSize; j++)  //�������Ʋ���Ҫ ��1��ʼ
		{
			CString sTempTime = vTempRawData[j];
			sTempTime.Trim();
			if (j >= 5)
			{
				//��5�п�ʼ ��ʱ�����
				switch (BelongToWhichTime(sTempTime))
				{
					case MORNING:
						if (vTempMorningTime.size() < 2)   // ����2�δ�ʱ��ֻ��ǰ����
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
				// 1-4��ԭ�����
				vTempResultData.push_back(sTempTime);
			}
		}

		// ���硢�������
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
		sTempString.Format(L"%d����", iMorningTime);
		vTempResultData.push_back(sTempString);
		sTempString.Format(L"%d����", iAfternoonTime);
		vTempResultData.push_back(sTempString);
		sTempString.Format(L"%d����", iEveningTime);
		vTempResultData.push_back(sTempString);
		sTempString.Format(L"%d����", iAllDayTime);
		vTempResultData.push_back(sTempString);

		// ��ÿ�н������
		vResultDatas.push_back(vTempResultData);
		if (sWeekTime == "������") // ÿλͬѧһ����������¼��������ʱ��
		{
			vector<CString> vSumTimes;
			sTempString.Format(L"�ܷ���:%d����  ��Сʱ:%dСʱ", iSumMinute, iSumMinute/60);
			vSumTimes.push_back(sTempString);
			vResultDatas.push_back(vSumTimes);

			vector<CString> vTempSummary;
			CString sPersonNum = vTempResultData[0];  // ���±��
			CString sPersonName = vTempResultData[1]; // ����
			sTempString.Format(L"%d", iSumMinute/60);
			vTempSummary.push_back(sPersonNum);
			vTempSummary.push_back(sPersonName);
			vTempSummary.push_back(sTempString);
			vSummaryTables.push_back(vTempSummary);

			iSumMinute = 0;
		}
	}


	//����Excel�ļ�
	CDatabase DB;

	//Excel��װ����
	CString StrDriver = "MICROSOFT EXCEL DRIVER (*.XLS)";

	//Ҫ������Execel�ļ�
	CString StrExcelFile = sPath;
	CString StrSQL;
	StrSQL.Format(L"DRIVER={%s};DSN='';FIRSTROWHASNameS=1;READONLY=FALSE;CREATE_DB=%s;DBQ=%s", StrDriver, StrExcelFile, StrExcelFile);
	TRY
	{
		//����Excel����ļ�
		DB.OpenEx(StrSQL, CDatabase::noOdbcDialog);

		//������ṹ���ֶ���������Index
		StrSQL = "CREATE TABLE ѧ��������ϸ��(���±�� TEXT, ���� TEXT, ˢ������ TEXT, ���� TEXT, ������� TEXT, ������� TEXT, ����ʱ�� TEXT, ����ʱ�� TEXT, ����ʱ�� TEXT, һ����ʱ�� TEXT)";
		DB.ExecuteSQL(StrSQL);

		//����ѧ��������ϸ������
		int iResultDatasSize = vResultDatas.size();
		for (int i = 0; i < iResultDatasSize; i++)
		{
			vector<CString> vTempRawData = vResultDatas[i];
			if (10 == vTempRawData.size())
			{
				StrSQL.Format(L"INSERT INTO ѧ��������ϸ�� (���±��, ����, ˢ������, ����, �������, �������, ����ʱ��, ����ʱ��, ����ʱ��, һ����ʱ��) VALUES ('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')", vTempRawData[0], vTempRawData[1], vTempRawData[2], vTempRawData[3], vTempRawData[4], vTempRawData[5], vTempRawData[6], vTempRawData[7], vTempRawData[8], vTempRawData[9]);
				DB.ExecuteSQL(StrSQL);
			}
			else
			{
				StrSQL.Format(L"INSERT INTO ѧ��������ϸ�� (���±��) VALUES ('%s')", vTempRawData[0]);
				DB.ExecuteSQL(StrSQL);
				if (i != iResultDatasSize - 1)
				{
					StrSQL.Format(L"INSERT INTO ѧ��������ϸ�� (���±��) VALUES ('')");
					DB.ExecuteSQL(StrSQL);
					StrSQL.Format(L"INSERT INTO ѧ��������ϸ�� (���±��, ����, ˢ������, ����, �������, �������, ����ʱ��, ����ʱ��, ����ʱ��, һ����ʱ��) VALUES ('���±��', '����', 'ˢ������', '����', '�������', '�������', '����ʱ��', '����ʱ��', '����ʱ��', 'һ����ʱ��')");
					DB.ExecuteSQL(StrSQL);
				}
			}
		}

		// ����CBIB�ܿ����ܽ�� 
		StrSQL = "CREATE TABLE CBIB�ܿ����ܽ��(���±�� TEXT, ���� TEXT, ��ʱ�� TEXT, ��ע TEXT)";
		DB.ExecuteSQL(StrSQL);
		COleDateTime startDate;
		COleDateTime endDate;
		m_startDateCtrl.GetTime(startDate);
		m_endDateCtrl.GetTime(endDate);
		CString sStartDateTime;
		CString sEndDateTime;
		sStartDateTime.Format(L"%d/%d/%d", startDate.GetYear(), startDate.GetMonth(), startDate.GetDay());
		sEndDateTime.Format(L"%d/%d/%d", endDate.GetYear(), endDate.GetMonth(), endDate.GetDay());
		StrSQL.Format(L"INSERT INTO CBIB�ܿ����ܽ�� (���±��) VALUES ('CBIB�ܿ����ܽ�� %s - %s')", sStartDateTime, sEndDateTime);
		DB.ExecuteSQL(StrSQL);

		int iSummaryTablesSize = vSummaryTables.size();
		for (int i = 0; i < iSummaryTablesSize; i++)
		{
			vector<CString> vTempRawData = vSummaryTables[i];
			StrSQL.Format(L"INSERT INTO CBIB�ܿ����ܽ�� (���±��, ����, ��ʱ��, ��ע) VALUES ('%s', '%s', '%s', '')", vTempRawData[0], vTempRawData[1], vTempRawData[2]);
			DB.ExecuteSQL(StrSQL);
		}


		//�ر����ݿ�
		DB.Close();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox("Excel��񴴽�����" + e->m_strError);
		return;
	}
	END_CATCH;

	MessageBox(L"�����ɹ���");
}


void CReadExcelDlg::OnClickedButtonSaveIni()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CFileFind fileFinder;   //�����Ƿ����ini�ļ�
	BOOL isFind = fileFinder.FindFile(_T(".\\AttendanceStatistic.ini"));
	if (isFind)
	{
		if (::MessageBox(NULL, L"ȷ��Ҫ����ԭ�е�������", L"��������", MB_YESNO) == IDNO)
		{
			return;
		}
	}

	// ��Ч��ʱ������
	
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningStart"), m_TimeMorningStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeMorningEnd"), m_TimeMorningEnd, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonStart"), m_TimeAfternoonStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeAfternoonEnd"), m_TimeAfternoonEnd, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningStart"), m_TimeEveningStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Effective Time"), _T("TimeEveningEnd"), m_TimeEveningEnd, _T(".\\AttendanceStatistic.ini"));

	// ����ʱ������
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningStart"), m_WorkTimeMorningStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeMorningEnd"), m_WorkTimeMorningEnd, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonStart"), m_WorkTimeAfternoonStart, _T(".\\AttendanceStatistic.ini"));
	::WritePrivateProfileStringW(_T("Work Time"), _T("WorkTimeAfternoonEnd"), m_WorkTimeAfternoonEnd, _T(".\\AttendanceStatistic.ini"));
	MessageBox(L"���ñ���ɹ�!");
	fileFinder.Close();

	// ȫ�����³�ʼ��
	vOriDatas.clear();
	vResultDatas.clear();
	vSummaryTables.clear();
	GetDlgItem(IDC_ExportExcel)->EnableWindow(FALSE);
	SetDlgItemText(IDC_EDIT_FILEPATH, L"");
}

void CReadExcelDlg::OnChangeEditMorningStart()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_EDIT_MORNING_START, m_TimeMorningStart);
}


void CReadExcelDlg::OnChangeEditMorningEnd()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_EDIT_MORNING_END, m_TimeMorningEnd);
}


void CReadExcelDlg::OnChangeEditAfternoonStart()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_EDIT_AFTERNOON_START, m_TimeAfternoonStart);
}


void CReadExcelDlg::OnChangeEditAfternoonEnd()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_EDIT_AFTERNOON_END, m_TimeAfternoonEnd);
}


void CReadExcelDlg::OnChangeEditEveningStart()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_EDIT_EVENING_START, m_TimeEveningStart);
}


void CReadExcelDlg::OnChangeEditEveningEnd()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_EDIT_EVENING_END, m_TimeEveningEnd);
}


void CReadExcelDlg::OnChangeWorktimeEditMorningStart()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_WORKTIME_EDIT_MORNING_START, m_WorkTimeMorningStart);
}


void CReadExcelDlg::OnChangeWorktimeEditMorningEnd()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_WORKTIME_EDIT_MORNING_END, m_WorkTimeMorningEnd);
}


void CReadExcelDlg::OnChangeWorktimeEditAfternoonStart()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_START, m_WorkTimeAfternoonStart);
}


void CReadExcelDlg::OnChangeWorktimeEditAfternoonEnd()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	GetDlgItemText(IDC_WORKTIME_EDIT_AFTERNOON_END, m_WorkTimeAfternoonEnd);
}
