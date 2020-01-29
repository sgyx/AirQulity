
// AirFQualityDlg.h: 头文件
//

#pragma once


// CAirFQualityDlg 对话框
class CAirFQualityDlg : public CDialogEx
{
// 构造
public:
	CAirFQualityDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_AIRFQUALITY_DIALOG };
#endif

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
	afx_msg CString SelFilePath();
	DECLARE_MESSAGE_MAP();
public:
	afx_msg void OnEnChangeEdit1();
	afx_msg void OnBnClickedbotton2();
	afx_msg void OnBnClickedbotton1();
	BOOL ExcuteCmd(CString bstrSqlCmd);
	afx_msg void OnBnClickedbotton3();
	afx_msg void OnBnClickedbotton4();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedOk();
};
