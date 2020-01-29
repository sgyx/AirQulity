
// AirFQualityDlg.cpp: 实现文件
//

#include "stdafx.h"
#include "AirFQuality.h"
#include "AirFQualityDlg.h"
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
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CAirFQualityDlg 对话框



CAirFQualityDlg::CAirFQualityDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_AIRFQUALITY_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CAirFQualityDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAirFQualityDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//ON_EN_CHANGE(IDC_EDIT1, &CAirFQualityDlg::OnEnChangeEdit1)
	ON_BN_CLICKED(botton2, &CAirFQualityDlg::OnBnClickedbotton2)
	ON_BN_CLICKED(botton1, &CAirFQualityDlg::OnBnClickedbotton1)
	ON_BN_CLICKED(botton3, &CAirFQualityDlg::OnBnClickedbotton3)
	ON_BN_CLICKED(botton4, &CAirFQualityDlg::OnBnClickedbotton4)
	//ON_BN_CLICKED(IDC_BUTTON1, &CAirFQualityDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDOK, &CAirFQualityDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CAirFQualityDlg 消息处理程序

BOOL CAirFQualityDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
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

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CAirFQualityDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CAirFQualityDlg::OnPaint()
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
HCURSOR CAirFQualityDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CAirFQualityDlg::OnEnChangeEdit1()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}

_ConnectionPtr  m_pConnection;		//连接对象
_RecordsetPtr  m_pRecordSet;		//记录集对象


BOOL CAirFQualityDlg::ExcuteCmd(CString bstrSqlCmd)//执行SQL语句
{
	CAirFQualityDlg ct;
	_bstr_t m_sqlCmd = _bstr_t(bstrSqlCmd);
	_variant_t RecordsAffected;

	try {
		m_pRecordSet = m_pConnection->Execute(m_sqlCmd, &RecordsAffected, adCmdText);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
		return FALSE;
	}
	return TRUE;
}


//输入1个字节输出对应的十进制整数
int Calculated_data(LPTSTR b) {
	int x = 0;
	int y = 0;
	if (b[0] >= '0' && b[0] <= '9') x = b[0] - '0';
	if (b[0] >= 'a' && b[0] <= 'f') x = b[0] - 'a' + 10;
	if (b[0] >= 'A' && b[0] <= 'F') x = b[0] - 'A' + 10;
	if (b[1] >= '0' && b[1] <= '9') y = b[1] - '0';
	if (b[1] >= 'a' && b[1] <= 'f') y = b[1] - 'a' + 10;
	if (b[1] >= 'A' && b[1] <= 'F') y = b[1] - 'A' + 10;
	return x * 16 + y;
}

void CAirFQualityDlg::OnBnClickedbotton2()
{
	
	CStdioFile file;
	CString AllData, Check1, Check2, CurrentData, HandleData;
	CString out;
	char *cAllchar;
	CString Data[7];
	CString str = _T("\r\n");
	CEdit* pdisplay1 = (CEdit*)GetDlgItem(display1);
	CEdit* pdisplay2 = (CEdit*)GetDlgItem(display2);
	CEdit* pdisplay3 = (CEdit*)GetDlgItem(display3);
	CEdit* pdisplay4 = (CEdit*)GetDlgItem(display4);
	CEdit* pdisplay5 = (CEdit*)GetDlgItem(display5);
	CEdit* pdisplay6 = (CEdit*)GetDlgItem(display6);
	CEdit* pdisplay7 = (CEdit*)GetDlgItem(display7);
	SYSTEMTIME sys;
	CTime systim = CTime::GetCurrentTime();
	long long flag;
	int use;
	int sum = 0;
	int count;
	int temporary = 0;
	int nNum1;
	float fNum2;
	CString strFile = _T("usetext.txt");
	CFileDialog    dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY, _T("Describe Files (*.txt)|*.txt|All Files (*.*)|*.*||"), NULL);
	if (dlgFile.DoModal()) {
		strFile = dlgFile.GetPathName();
	}

	if (!(file.Open(strFile, CFile::typeBinary))) {         //打开文件并读取其中的值
		MessageBox(_T("Failed to open the file!"));
		return;
	}
	
	file.SeekToBegin();
	//file.ReadString(AllData);
	int nFileLength;
	nFileLength = file.GetLength();
	BYTE *tesss;
	tesss = (BYTE *)malloc(sizeof(BYTE) * nFileLength);
	file.Read(tesss, nFileLength);
	for (flag = 0; flag <= nFileLength; flag += 17) {            //循环每一串数据
		/*CurrentData = AllData.Mid(flag, 34);
		Check1 = CurrentData.Mid(0, 2);
		Check2 = CurrentData.Mid(2, 2);*/
		if (tesss[flag] == 0x3C && tesss[flag + 1] == 0x02) {
			int sum = 0;
				for (int test = 0; test <= 15; test++) {
					sum += tesss[flag + test];
				}
				if ((int)tesss[flag + 16] == (int)(sum & 255)) {
					LPTSTR a = (LPTSTR)(LPCTSTR)CurrentData;						//输出数据
						CString Sql;												//输出到数据库
						CString oco2, och2o, otvoc, opm2, opm10, otemperature, ohumidity;
						CString oyear, omonth, oday, ohour, ominute, osecond, omillisecond;
						
						CString odeafule;
						int nOco2, nOch2o, nOtvoc, nOpm2, nOpm10;
						float fOtemperature, fOhumidity;
						GetLocalTime(&sys);
						nOco2 = tesss[flag+2]*255 + tesss[flag + 3];
						nOch2o = tesss[flag + 4] * 256 + tesss[flag + 5];
						nOtvoc = tesss[flag + 6] * 256 + tesss[flag + 7];
						nOpm2 = tesss[flag + 8] * 256 + tesss[flag + 9];
						nOpm10 = tesss[flag + 10] * 256 + tesss[flag + 11];
						fOtemperature = tesss[flag + 12] + tesss[flag + 13] * 0.1;
						fOhumidity = tesss[flag + 14] + tesss[flag + 15] * 0.1;

						oco2.Format(_T("%d"), nOco2);
						och2o.Format(_T("%d"), nOch2o);
						otvoc.Format(_T("%d"), nOtvoc);
						opm2.Format(_T("%d"), nOpm2);
						opm10.Format(_T("%d"), nOpm10);
						otemperature.Format(_T("%.2f"), fOtemperature);
						ohumidity.Format(_T("%.2f"), fOhumidity);

						systim += CTimeSpan(0, 0, 0, 1);
						systim.GetAsSystemTime(sys);

						oyear.Format(_T("%4d"), sys.wYear);
						omonth.Format(_T("%2d"), sys.wMonth);
						oday.Format(_T("%2d"), sys.wDay);
						ohour.Format(_T("%2d"), sys.wHour);
						ominute.Format(_T("%2d"), sys.wMinute);
						osecond.Format(_T("%2d"), sys.wSecond);
						omillisecond.Format(_T("%3d"), sys.wMilliseconds);
						
						odeafule = _T("insert into air(Atime, co2, ch2o, tvoc, pm2, pm10, temperature, humidity) values('");

						Sql = odeafule + oyear + _T("-") + omonth + _T("-") + oday + _T(" ") + ohour + _T(":") 
							+ ominute + _T(":") + osecond + _T(".") + omillisecond + _T("',")
							+ oco2 + _T(",") + och2o + _T(",") + otvoc + _T(",")
							+ opm2 + _T(",") + opm10 + _T(",")
							+ otemperature + _T(",") + ohumidity + _T(")");
						//OnInitDialog();
						if (!ExcuteCmd(Sql))
							AfxMessageBox(_T("添加失败！"));
						/*else
							AfxMessageBox(_T("添加成功！"));*/
	
						int nLenText = pdisplay1->GetWindowTextLength();					//输出到屏幕
						out.Format(_T("%d"), nOco2);
						pdisplay1->SetSel(nLenText, nLenText);
						pdisplay1->ReplaceSel(out + str);

						out.Format(_T("%d"), nOch2o);
						pdisplay2->SetSel(-1, -1, true);
						pdisplay2->ReplaceSel(out + str);


						out.Format(_T("%d"), nOtvoc);
						pdisplay3->SetSel(-1, -1, true);
						pdisplay3->ReplaceSel(out + str);

						out.Format(_T("%d"), nOpm2);
						pdisplay4->SetSel(-1, -1, true);
						pdisplay4->ReplaceSel(out + str);

						out.Format(_T("%d"), nOpm10);
						pdisplay5->SetSel(-1, -1, true);
						pdisplay5->ReplaceSel(out + str);

						out.Format(_T("%.2f"), fOtemperature);
						pdisplay6->SetSel(-1, -1, true);
						pdisplay6->ReplaceSel(out + str);

						out.Format(_T("%.2f"), fOhumidity);
						pdisplay7->SetSel(-1, -1, true);
						pdisplay7->ReplaceSel(out + str);
			}
		}
	}
	free(tesss);
	AfxMessageBox(_T("写入完成！"));
}

CString CAirFQualityDlg::SelFilePath()
{
	TCHAR szFolderPath[MAX_PATH] = { 0 };
	CString strFolderPath = TEXT("");
	BROWSEINFO      sInfo;
	::ZeroMemory(&sInfo, sizeof(BROWSEINFO));
	sInfo.pidlRoot = 0;
	sInfo.lpszTitle = _T("请选择处理结果存储路径");
	sInfo.ulFlags = BIF_RETURNONLYFSDIRS | BIF_EDITBOX | BIF_DONTGOBELOWDOMAIN;
	sInfo.lpfn = NULL;

	// 显示文件夹选择对话框  
	LPITEMIDLIST lpidlBrowse = ::SHBrowseForFolder(&sInfo);
	if (lpidlBrowse != NULL) {
		// 取得文件夹名  
		if (::SHGetPathFromIDList(lpidlBrowse, szFolderPath))
		{
			strFolderPath = szFolderPath;
		}
	}
	if (lpidlBrowse != NULL) {
		CoTaskMemFree(lpidlBrowse);

		return strFolderPath;
	}
}

void CAirFQualityDlg::OnBnClickedbotton1()
{
	AfxOleInit();
	HRESULT hr;
	hr = m_pConnection.CreateInstance("ADODB.Connection");
	m_pConnection->ConnectionTimeout = 8;
	try
	{
		_bstr_t strConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=test;Password=20001101;Initial Catalog=airqulity;Data Source=LAPTOP-VGEKT9PL";//用户名及密码  
		hr = m_pConnection->Open(strConnect, "", "", adModeUnknown);//打开数据库
		if (FAILED(hr))
		{
			AfxMessageBox(_T("不能连接数据库!"));
			return;
		}
		else
			AfxMessageBox(_T("连接数据库成功!"));
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());

		return;
	}
}




void CAirFQualityDlg::OnBnClickedbotton3()
{
	CString st1;
	CStdioFile output;
	st1 = _T("select * from air");
	CString strFile = SelFilePath() + _T(".txt");
	//CString strFile = _T("AirQualityOut.txt");
	/*CFileDialog    dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY, _T("Describe Files (*.txt)|*.txt|All Files (*.*)|*.*||"), NULL);
	if (dlgFile.DoModal()){
		strFile = dlgFile.GetPathName();
	}*/
	if (!output.Open(strFile, CFile::modeWrite | CFile::modeCreate | CFile::typeText)) {
		MessageBox(_T("Fail Reading Failure!"));
		return;
	}
	if (!ExcuteCmd(st1))
		AfxMessageBox(_T("执行失败！"));
	else
		AfxMessageBox(_T("执行成功！"));
	while (!m_pRecordSet -> adoEOF) {
		output.WriteString((CString)m_pRecordSet->GetCollect("Atime") + _T(" ")
			+ (CString)m_pRecordSet->GetCollect("co2") + _T(" ")
			+ (CString)m_pRecordSet->GetCollect("ch2o") + _T(" ")
			+ (CString)m_pRecordSet->GetCollect("tvoc") + _T(" ")
			+ (CString)m_pRecordSet->GetCollect("pm2") + _T(" ")
			+ (CString)m_pRecordSet->GetCollect("pm10") + _T(" ")
			+ (CString)m_pRecordSet->GetCollect("temperature") + _T(" ")
			+ (CString)m_pRecordSet->GetCollect("humidity") + _T("\n"));
		m_pRecordSet->MoveNext();
	}

	AfxMessageBox(_T("写入成功！"));
}


void CAirFQualityDlg::OnBnClickedbotton4()
{
	CString st1;
	CString a, b;
	CListCtrl* mygrid = (CListCtrl*)GetDlgItem(reporttest);
	st1 = _T("select * from air");
	if (!ExcuteCmd(st1))
		AfxMessageBox(_T("执行失败！"));
	else
		AfxMessageBox(_T("执行成功！"));
	m_pRecordSet->MoveFirst();
	mygrid->InsertColumn(0, _T("TIME"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(1, _T("CO2"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(2, _T("CH2O"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(3, _T("TVOC"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(4, _T("PM2.5"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(5, _T("PM10"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(6, _T("TEMPERATURE"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(7, _T("HUMIDILY"), LVCFMT_CENTER, 95);
	while (!m_pRecordSet->adoEOF) {
		int flag = 0;
		mygrid->InsertItem(flag, _T(" "));
		mygrid->SetItemText(flag, 0, (CString)m_pRecordSet->GetCollect("Atime"));
		mygrid->SetItemText(flag, 1, (CString)m_pRecordSet->GetCollect("co2"));
		mygrid->SetItemText(flag, 2, (CString)m_pRecordSet->GetCollect("ch2o"));
		mygrid->SetItemText(flag, 3, (CString)m_pRecordSet->GetCollect("tvoc"));
		mygrid->SetItemText(flag, 4, (CString)m_pRecordSet->GetCollect("pm2"));
		mygrid->SetItemText(flag, 5, (CString)m_pRecordSet->GetCollect("pm10"));
		mygrid->SetItemText(flag, 6, (CString)m_pRecordSet->GetCollect("temperature"));
		mygrid->SetItemText(flag, 7, (CString)m_pRecordSet->GetCollect("humidity"));
		flag++;
		m_pRecordSet->MoveNext();
	}
}


void CAirFQualityDlg::OnBnClickedButton1()
{
	CListCtrl* mygrid = (CListCtrl*)GetDlgItem(reporttest);
	mygrid->InsertColumn(0, _T("TIME"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(1, _T("CO2"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(2, _T("CH2O"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(3, _T("TVOC"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(4, _T("PM2.5"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(5, _T("PM10"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(6, _T("TEMPERATURE"), LVCFMT_CENTER, 95);
	mygrid->InsertColumn(7, _T("HUMIDILY"), LVCFMT_CENTER, 95);
	mygrid->InsertItem(0,_T(" "));
	//mygrid->InsertItem(0, _T("111"));
	mygrid->SetItemText(0, 0, _T("1111"));
}


void CAirFQualityDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
}
