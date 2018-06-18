
// MFC_OPC_ExplorerDlg.cpp : файл реализации
//

#include "stdafx.h"
#include "MFC_OPC_Explorer.h"
#include "MFC_OPC_ExplorerDlg.h"
#include "afxdialogex.h"

// error LNK2001: неразрешенный внешний символ "_IID_IOPCServer"
// error LNK2001: неразрешенный внешний символ "_IID_IOPCItemMgt"
// Захотим вызвать ошибку закомментим include и раскоментим два const
#include "opcda_i.c"

//#include "tree.h"
//#include "multitree.h"
//#include "unique_tree.h"
//#include "sequential_tree.h"

#include <Psapi.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Диалоговое окно CAboutDlg используется для описания сведений о приложении

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Данные диалогового окна
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // поддержка DDX/DDV

// Реализация
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


// диалоговое окно CMFC_OPC_ExplorerDlg




CMFC_OPC_ExplorerDlg::CMFC_OPC_ExplorerDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CMFC_OPC_ExplorerDlg::IDD, pParent)
	, m_BtnTagReadFromCahche(0)
	, m_BtnTagReadFromDevice(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMFC_OPC_ExplorerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_HOST, m_EditHost);
	DDX_Control(pDX, IDC_EDIT_PROGID, m_EditProgID);
	DDX_Control(pDX, IDC_EDIT_TIMEOUT_CONNECT, m_EditTimeoutConnect);
	DDX_Control(pDX, IDC_EDIT_TIMEOUT_TAGS, m_EditTimeoutTags);
	DDX_Control(pDX, IDC_BTN_CONNECT, m_BtnConnect);
	DDX_Control(pDX, IDC_STATIC_CONN_STATUS, m_StaticConnectStatus);
	DDX_Control(pDX, IDC_BTN_DISCONNECT, m_BtnDisconnect);
	DDX_Control(pDX, IDC_LIST_CONNECT_STATUS, m_ListConnect);
	//	DDX_Control(pDX, IDC_BTN_CONNECT_REFRESH, m_BtnConnectRefresh);
	DDX_Control(pDX, IDC_EDIT_GROUP_TAG_NAME, m_EditGroupTagName);
	DDX_Control(pDX, IDC_BTN_GROUP_TAG_NAME, m_BtnGroupTagName);
	DDX_Control(pDX, IDC_EDIT_TAG_NAME, m_EditTagName);
	DDX_Control(pDX, IDC_CHECK_TAG_TYPE, m_CheckTagType);
	DDX_Control(pDX, IDC_CMBX_TAG_TYPE, m_CmBxTagType);
	DDX_Control(pDX, IDC_EDIT_TAG_TYPE, m_EditTagType);
	DDX_Control(pDX, IDC_EDIT_TAG_VALUE, m_EditTagValue);
	DDX_Control(pDX, IDC_EDIT_TAG_QUALITY, m_StaticTagQual);
	DDX_Control(pDX, IDC_CHECK_TAG_READ_WRITE, m_CheckTagReadWrite);
	DDX_Control(pDX, IDC_BTN_TAG_READ_WRITE, m_BtnTagReadWrite);
	DDX_Control(pDX, IDC_LIST_TAG_READ_WRITE_STATUS, m_ListTagReadWrite);
	//	DDX_Control(pDX, IDC_BTN_TAG_READ_WRITE_REFRESH, m_LstBxTagReadWriteRefresh);
	DDX_Control(pDX, IDC_BTN_GET_SERVER_ADDRESS_SPACE, m_BtnGetSrvrAddrSpace);
	DDX_Control(pDX, IDC_TREE_SERVER_ADDRESS_SPACE, m_TreeGetSrvrAddrSpace);
	DDX_Control(pDX, IDC_LIST_SERVER_ADDRESS_SPACE_STATUS, m_ListGetSrvrAddrSpace);
	//	DDX_Control(pDX, IDC_BTN_SERVER_ADDRESS_SPACE_REFRESH, m_LstBxGetSrvrAddrSpaceRefresh);
	DDX_Control(pDX, IDC_LIST_TAG_PROPERTIES, m_LstBxTagProperties);
	DDX_Control(pDX, IDC_LIST_TAG_RW_PROPERTIES, m_LstBxTagRWProperties);
}

BEGIN_MESSAGE_MAP(CMFC_OPC_ExplorerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CMFC_OPC_ExplorerDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CMFC_OPC_ExplorerDlg::OnBnClickedCancel)
	ON_EN_CHANGE(IDC_EDIT_HOST, &CMFC_OPC_ExplorerDlg::OnEnChangeEditHost)
	ON_EN_CHANGE(IDC_EDIT_PROGID, &CMFC_OPC_ExplorerDlg::OnEnChangeEditProgid)
	ON_EN_CHANGE(IDC_EDIT_TIMEOUT_CONNECT, &CMFC_OPC_ExplorerDlg::OnEnChangeEditTimeoutConnect)
	ON_EN_CHANGE(IDC_EDIT_TIMEOUT_TAGS, &CMFC_OPC_ExplorerDlg::OnEnChangeEditTimeoutTags)
	ON_BN_CLICKED(IDC_BTN_CONNECT, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnConnect)
	ON_BN_CLICKED(IDC_BTN_DISCONNECT, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnDisconnect)
	ON_BN_CLICKED(IDC_BTN_CONNECT_REFRESH, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnConnectRefresh)
	ON_EN_CHANGE(IDC_EDIT_GROUP_TAG_NAME, &CMFC_OPC_ExplorerDlg::OnEnChangeEditGroupTagName)
	ON_BN_CLICKED(IDC_BTN_GROUP_TAG_NAME, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnGroupTagName)
	ON_EN_CHANGE(IDC_EDIT_TAG_NAME, &CMFC_OPC_ExplorerDlg::OnEnChangeEditTagName)
	ON_BN_CLICKED(IDC_CHECK_TAG_TYPE, &CMFC_OPC_ExplorerDlg::OnBnClickedCheckTagType)
	ON_CBN_SELCHANGE(IDC_CMBX_TAG_TYPE, &CMFC_OPC_ExplorerDlg::OnCbnSelchangeCmbxTagType)
	ON_EN_CHANGE(IDC_EDIT_TAG_TYPE, &CMFC_OPC_ExplorerDlg::OnEnChangeEditTagType)
	ON_EN_CHANGE(IDC_EDIT_TAG_VALUE, &CMFC_OPC_ExplorerDlg::OnEnChangeEditTagValue)
	ON_BN_CLICKED(IDC_CHECK_TAG_READ_WRITE, &CMFC_OPC_ExplorerDlg::OnBnClickedCheckTagReadWrite)
	ON_BN_CLICKED(IDC_BTN_TAG_READ_WRITE, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadWrite)
	ON_BN_CLICKED(IDC_BTN_TAG_READ_WRITE_REFRESH, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadWriteRefresh)
	ON_BN_CLICKED(IDC_BTN_GET_SERVER_ADDRESS_SPACE, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnGetServerAddressSpace)
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE_SERVER_ADDRESS_SPACE, &CMFC_OPC_ExplorerDlg::OnTvnSelchangedTreeServerAddressSpace)
	ON_BN_CLICKED(IDC_BTN_SERVER_ADDRESS_SPACE_REFRESH, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnServerAddressSpaceRefresh)
	ON_BN_CLICKED(IDC_BTN_TAG_READ_CACHE, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadCache)
	ON_BN_CLICKED(IDC_BTN_TAG_READ_DEVICE, &CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadDevice)
END_MESSAGE_MAP()


// обработчики сообщений CMFC_OPC_ExplorerDlg

BOOL CMFC_OPC_ExplorerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Добавление пункта "О программе..." в системное меню.

	// IDM_ABOUTBOX должен быть в пределах системной команды.
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

	// Задает значок для этого диалогового окна. Среда делает это автоматически,
	//  если главное окно приложения не является диалоговым
	SetIcon(m_hIcon, TRUE);			// Крупный значок
	SetIcon(m_hIcon, FALSE);		// Мелкий значок

	// TODO: добавьте дополнительную инициализацию

	/*
	//CString sTagsCSV = _T("123.txt");
	//CString sTagsCSV = _T("Теги-Номера-Задвижек.csv");
	CString sTagsCSV = _T("Теги-Номера-Задвижек-NewLine.csv");
	
	FILE *fStream;
	//errno_t e = _tfopen_s(&fStream, sTagsCSV, _T("rt,ccs=UNICODE"));
	errno_t e = _tfopen_s(&fStream, sTagsCSV, _T("rt,ccs=UTF-8"));
	
	if(e == 0) ; // failed..CString sRead;
	{
		CStdioFile f(fStream);  // open the file from this stream
	
		CString sRead;
		while(true)
		{
			//line_CStr.Empty();
			if(!f.ReadString(sRead))
			{break;}
			else
			{
				int a = sRead.Find(_T("\n"));
				int b = sRead.Find(_T("|"));
				int u(0);
			}
		}
	

		f.Close();
	}

	*/

	//COLORREF a(RGB(1,1,1));
	//m_TreeGetSrvrAddrSpace.SetBkColor(COLORREF(RGB(0,0,255))); // BLUE
	m_TreeGetSrvrAddrSpace.SetBkColor(COLORREF(RGB(220,255,220))); // wHITE_GREEN
	
/*////////////////////////////////////////////////////////////////////////
	// 1 // Получаем путь к exe-файлу приложения и имя ini-файла сопряжённого с ним
	TCHAR IniFileName[_MAX_PATH];
	TCHAR IniFilePath[_MAX_PATH];
	::GetModuleFileName(AfxGetApp()->m_hInstance, IniFileName, MAX_PATH);
	for(int i=0; i<_MAX_PATH; i++) {IniFilePath[i] = IniFileName[i];}

	TCHAR* szExt = _tcsrchr(IniFileName, '.');
	if(szExt) *(szExt+1) = 0;
	szExt = _tcsrchr(IniFilePath, '\\');
	if(szExt) *(szExt+1) = 0;
	//приставить расширение
	_tcscat_s(IniFileName, _MAX_PATH, _T("ini"));

	theApp.sIniFileName = IniFileName; //Путь до ini-файла, включающий имя самого файла в этом пути
	theApp.sIniFilePath = IniFilePath;
	*/

//*///////////////////////////////////////////////////////////////////////
	// 2 // Получаем путь к exe-файлу приложения и имя ini-файла сопряжённого с ним
	//CHAR szFullModulePath[_MAX_PATH]; szFullModulePath[0]=0;
	TCHAR szFullModulePath[_MAX_PATH]; szFullModulePath[0]=0;
	//TCHAR IniFileName[_MAX_PATH];
	TCHAR IniFilePath[_MAX_PATH];

	HANDLE hProcess = ::OpenProcess(
									PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
									FALSE,
									GetCurrentProcessId()
									);
	if(NULL != hProcess)
	{
		//if( ::GetModuleFileNameExA( hProcess, NULL, szFullModulePath, _MAX_PATH ) )
		if( ::GetModuleFileNameExW( hProcess, NULL, szFullModulePath, _MAX_PATH ) )
		{
			for(int i=0; i<_MAX_PATH; i++) {IniFilePath[i] = szFullModulePath[i];}

			TCHAR* szExt = _tcsrchr(szFullModulePath, '.');
			if(szExt) *(szExt+1) = 0;
			//отрезать имя
			//CHAR* szName = strrchr(szFullModulePath, '\\');
			TCHAR* szName = _tcsrchr(IniFilePath, '\\');
			if(szName) *(szName+1) = 0;
			//приставить расширение
			_tcscat_s(szFullModulePath, _MAX_PATH, _T("ini"));
			//вставить путь
			//m_prot_name.insert(0,szFullModulePath);

		}
		CloseHandle(hProcess);
	}

	theApp.sIniFileName = szFullModulePath; //Путь до ini-файла, включающий имя самого файла в этом пути
	theApp.sIniFilePath = IniFilePath;

	/////////////////////////////////////////////////////////////////////////
//*/


	// Получаем имя сервера (Хост) и идентификатор ОРС-сервера (ProgID)
	CIniParam parOPCServerParamsDescription(theApp.sIniFileName, _T("OPCServer"), _T("Description"), _T("0"));
	CIniParam parOPCServerHost(theApp.sIniFileName, _T("OPCServer"), _T("Host"), _T("localhost"));
	CIniParam parOPCServerProgID(theApp.sIniFileName, _T("OPCServer"), _T("ProgID"), _T("SymbolDesign.WatchSOU.1"));
	CIniParam parOPCServerConnTime(theApp.sIniFileName, _T("OPCServer"), _T("ConnTime"), _T("0.5"));
	CIniParam parOPCServerTagsTimeout(theApp.sIniFileName, _T("OPCServer"), _T("TagsTimeout"), _T("60"));

	CIniParam parOPCServerGroupTagName(theApp.sIniFileName, _T("OPCServer"), _T("GroupTagName"), _T("Grooppa_1"));
	CIniParam parOPCServerTagName(theApp.sIniFileName, _T("OPCServer"), _T("TagName"), _T("Pressure/77"));
	CIniParam parOPCServerTagType(theApp.sIniFileName, _T("OPCServer"), _T("TagType"), _T("VT_R4"));
	CIniParam parOPCServerTagValue(theApp.sIniFileName, _T("OPCServer"), _T("TagValue"), _T("0.789"));

	// Дополнительно в ini-шку заводим парамтр позволяющий посмотреть описание параметров, чтобы судорожно не вспоминать их потом.
	int nDescription = parOPCServerParamsDescription.AsInt();
	if(nDescription)
	{
		CString sDescription("Description");

		CIniParam descrOPCServerHost(theApp.sIniFileName,		 sDescription, _T("Host"),		 _T("имя удалённого сервера или его IP-адрес в явном виде"));
		CIniParam descrOPCServerProgID(theApp.sIniFileName,		 sDescription, _T("ProgID"),	 _T("имя ОРС сервера в системе"));
		CIniParam descrOPCServerConnTime(theApp.sIniFileName,	 sDescription, _T("ConnTime"),	 _T("время в секундах через сколько осуществлять переподключение к серверу"));
		CIniParam descrOPCServerTagsTimeout(theApp.sIniFileName, sDescription, _T("TagsTimeout"),_T("время в секундах сколько пробовать достучаться до тега"));
																 
		CIniParam descrOPCServerGroupTagName(theApp.sIniFileName,sDescription, _T("GroupTagName"),_T("имя группы для тегов, которое пропишется по умолчанию в интерфейсе при старте"));
		CIniParam descrOPCServerTagName(theApp.sIniFileName,	 sDescription, _T("TagName"),	 _T("имя тега, которое пропишется по умолчанию в интерфейсе при старте"));
		CIniParam descrOPCServerTagType(theApp.sIniFileName,	 sDescription, _T("TagType"),	 _T("тип тега, который пропишется по умолчанию в интерфейсе при старте"));
		CIniParam descrOPCServerTagValue(theApp.sIniFileName,	 sDescription, _T("TagValue"),	 _T("значение тега, которое пропишется по умолчанию в интерфейсе при старте"));
	}


	sOPCServerRmtSrvrNameOrIPAddr = parOPCServerHost.AsString();
	sOPCServerProgID = parOPCServerProgID.AsString();
	sOPCServerConnTime = parOPCServerConnTime.AsString();
	sOPCServerTagsTimeout = parOPCServerTagsTimeout.AsString();

	sGroupTagName = parOPCServerGroupTagName.AsString();
	sTagName = parOPCServerTagName.AsString();
	sTagType = parOPCServerTagType.AsString();
	sTagValue= parOPCServerTagValue.AsString();

	// в принципе можно оставить все 4 выражения ниже, что внутри if-ов без самих if-ов ....
	if(! sOPCServerRmtSrvrNameOrIPAddr.IsEmpty())	m_EditHost.SetWindowText(sOPCServerRmtSrvrNameOrIPAddr);
	if(! sOPCServerProgID.IsEmpty())				m_EditProgID.SetWindowText(sOPCServerProgID);
	if(! sOPCServerConnTime.IsEmpty())				m_EditTimeoutConnect.SetWindowText(sOPCServerConnTime);
	if(! sOPCServerTagsTimeout.IsEmpty())			m_EditTimeoutTags.SetWindowText(sOPCServerTagsTimeout);

	if(! sGroupTagName.IsEmpty())	m_EditGroupTagName.SetWindowText(sGroupTagName);
	if(! sTagName.IsEmpty())	m_EditTagName.SetWindowText(sTagName);
	if(! sTagType.IsEmpty())	m_EditTagType.SetWindowText(sTagType);
	if(! sTagValue.IsEmpty())	m_EditTagValue.SetWindowText(sTagValue);


	// Заносим в комбобокс все возможные типы
	SetCmBxTagTypes();


	CRect rect(NULL);
	m_ListConnect.GetClientRect(rect);
	/*
	CHeaderCtrl* pHeaderCtrl = m_ListConnect.GetHeaderCtrl(); //m_ListResult.GetHeaderCtrl();
	int nColumnCount = pHeaderCtrl->GetItemCount();
	m_ListConnect.RedrawItems(0,nColumnCount);
	m_ListConnect.DeleteAllItems();
	*/
	m_ListConnect.InsertColumn(0, _T("Этапы Подключения / Отключения"), LVCFMT_LEFT, 1000);

	m_ListTagReadWrite.GetClientRect(rect);
	m_ListTagReadWrite.InsertColumn(0, _T("Этапы Чтения / Записи тега"), LVCFMT_LEFT, 1000);

	m_ListGetSrvrAddrSpace.GetClientRect(rect);
	m_ListGetSrvrAddrSpace.InsertColumn(0, _T("Этапы получения адресного пространства"), LVCFMT_LEFT, 1000);


/*
	int curIndex=0;			// Индекс элемента выбранного в box
	sCurrStep = _T("Инициализация дескриптора окружения");
	m_ListConnect.InsertItem(m_ListConnect.GetItemCount(), sSuccses+sCurrStep);
//	m_ListConnect.RedrawItems(0,nColumnCount);
//	m_ListConnect.DeleteAllItems();

//	m_ListConnect.InsertColumn(0, _T("Статус подключения"), LVCFMT_LEFT, 1000);
	//m_ListConnect.InsertItem(0, _T("1"));
	//m_ListResult.SetItemText(i, j, v_SQLFullTable[i][j]);
*/

	// Помещаем в некоторые переменные их значения по умолчанию, ибо они должны быть устновлены при старте программы
	a_m_z_=0;
	bOPCConnection=false;
	bAddGroupTag=false;
	bReadWriteTag=false;
    UINT nCheckReadWrite = m_CheckTagReadWrite.GetState();

	this->CheckRadioButton(IDC_BTN_TAG_READ_CACHE, IDC_BTN_TAG_READ_DEVICE, IDC_BTN_TAG_READ_CACHE);
	//CButton* pButton = (CButton*)GetDlgItem(IDC_BTN_TAG_READ_CACHE);
	//pButton->SetCheck(true);


	// Обнулили Индексы для трёх наших ЛистБоксов
	//LstIndx.Prepare();

	
	return TRUE;  // возврат значения TRUE, если фокус не передан элементу управления
}

void CMFC_OPC_ExplorerDlg::SetCmBxTagTypes()
{
	m_CmBxTagType.ResetContent(); // удаляем всё из комбобокса
	// обрабатываем ситуацию в сответсвие с проставленной галкой
    UINT nCheck = m_CheckTagType.GetState();
    if(nCheck &= BST_CHECKED)
    {
		m_CmBxTagType.AddString( _T("VT_EMPTY")				);
		m_CmBxTagType.AddString( _T("VT_NULL")				);
		m_CmBxTagType.AddString( _T("VT_BOOL")				);
		m_CmBxTagType.AddString( _T("VT_I2")				);
		m_CmBxTagType.AddString( _T("VT_I4")				);
		m_CmBxTagType.AddString( _T("VT_R4")				);
		m_CmBxTagType.AddString( _T("VT_R8")				);
		m_CmBxTagType.AddString( _T("VT_BSTR")				);
		m_CmBxTagType.AddString( _T("VT_LPSTR")				);
		m_CmBxTagType.AddString( _T("VT_LPWSTR")			);
		m_CmBxTagType.AddString( _T("VT_DATE")				);
		m_CmBxTagType.AddString( _T("VT_ERROR")				);
		m_CmBxTagType.AddString( _T("VT_VARIANT")			);
		m_CmBxTagType.AddString( _T("VT_I1")				);
		m_CmBxTagType.AddString( _T("VT_UI1")				);
		m_CmBxTagType.AddString( _T("VT_UI2")				);
		m_CmBxTagType.AddString( _T("VT_UI4")				);
		m_CmBxTagType.AddString( _T("VT_I8")				);
		m_CmBxTagType.AddString( _T("VT_UI8")				);
		m_CmBxTagType.AddString( _T("VT_INT")				);
		m_CmBxTagType.AddString( _T("VT_UINT")				);
		m_CmBxTagType.AddString( _T("VT_INT_PTR")			);
		m_CmBxTagType.AddString( _T("VT_UINT_PTR")			);
		m_CmBxTagType.AddString( _T("VT_DECIMAL")			);
		m_CmBxTagType.AddString( _T("VT_HRESULT")			);
		m_CmBxTagType.AddString( _T("VT_VOID")				);
		m_CmBxTagType.AddString( _T("VT_CY")				);
		m_CmBxTagType.AddString( _T("VT_DISPATCH")			);
		m_CmBxTagType.AddString( _T("VT_UNKNOWN")			);
		m_CmBxTagType.AddString( _T("VT_RECORD")			);
		m_CmBxTagType.AddString( _T("VT_PTR")				);
		m_CmBxTagType.AddString( _T("VT_SAFEARRAY")			);
		m_CmBxTagType.AddString( _T("VT_CARRAY")			);
		m_CmBxTagType.AddString( _T("VT_USERDEFINED")		);
		m_CmBxTagType.AddString( _T("VT_FILETIME")			);
		m_CmBxTagType.AddString( _T("VT_BLOB")				);
		m_CmBxTagType.AddString( _T("VT_STREAM")			);
		m_CmBxTagType.AddString( _T("VT_STORAGE")			);
		m_CmBxTagType.AddString( _T("VT_STREAMED_OBJECT")	);
		m_CmBxTagType.AddString( _T("VT_STORED_OBJECT")		);
		m_CmBxTagType.AddString( _T("VT_VERSIONED_STREAM")	);
		m_CmBxTagType.AddString( _T("VT_BLOB_OBJECT")		);
		m_CmBxTagType.AddString( _T("VT_CF")				);
		m_CmBxTagType.AddString( _T("VT_CLSID")				);
		m_CmBxTagType.AddString( _T("VT_VECTOR")			);
		m_CmBxTagType.AddString( _T("VT_ARRAY")				);
		m_CmBxTagType.AddString( _T("VT_BYREF")				);
		m_CmBxTagType.AddString( _T("VT_BSTR_BLOB")			);
	}
    else //if(nCheck = BST_UNCHECKED)
    {
		m_CmBxTagType.AddString( _T("VT_EMPTY")				);
		m_CmBxTagType.AddString( _T("VT_NULL")				);
		m_CmBxTagType.AddString( _T("VT_BOOL")				);
		m_CmBxTagType.AddString( _T("VT_I2")				);
		m_CmBxTagType.AddString( _T("VT_I4")				);
		m_CmBxTagType.AddString( _T("VT_R4")				);
		m_CmBxTagType.AddString( _T("VT_R8")				);
		m_CmBxTagType.AddString( _T("VT_BSTR")				);
    }

}


void CMFC_OPC_ExplorerDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// При добавлении кнопки свертывания в диалоговое окно нужно воспользоваться приведенным ниже кодом,
//  чтобы нарисовать значок. Для приложений MFC, использующих модель документов или представлений,
//  это автоматически выполняется рабочей областью.

void CMFC_OPC_ExplorerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // контекст устройства для рисования

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Выравнивание значка по центру клиентского прямоугольника
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Нарисуйте значок
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// Система вызывает эту функцию для получения отображения курсора при перемещении
//  свернутого окна.
HCURSOR CMFC_OPC_ExplorerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//////////////////////////////////////////
// OK // CANCEL
void CMFC_OPC_ExplorerDlg::OnBnClickedOk()
{
	// TODO: добавьте свой код обработчика уведомлений
	CDialogEx::OnOK();
}


void CMFC_OPC_ExplorerDlg::OnBnClickedCancel()
{
	// TODO: добавьте свой код обработчика уведомлений
	CDialogEx::OnCancel();
}





////////////////////////////////////////////////////////////////
// 1 ///////////////////////////////////////////////////////////

void CMFC_OPC_ExplorerDlg::OnEnChangeEditHost()
{m_EditHost.GetWindowText(sOPCServerRmtSrvrNameOrIPAddr);}

void CMFC_OPC_ExplorerDlg::OnEnChangeEditProgid()
{m_EditProgID.GetWindowText(sOPCServerProgID);}

void CMFC_OPC_ExplorerDlg::OnEnChangeEditTimeoutConnect()
{m_EditTimeoutConnect.GetWindowText(sOPCServerConnTime);}

void CMFC_OPC_ExplorerDlg::OnEnChangeEditTimeoutTags()
{m_EditTimeoutTags.GetWindowText(sOPCServerTagsTimeout);}


void CMFC_OPC_ExplorerDlg::OnBnClickedBtnConnect()
{
	// TODO: добавьте свой код обработчика уведомлений
	AddItemToListBox(_T("Соединение..."), 1);
	bOPCConnection = Connect(sOPCServerRmtSrvrNameOrIPAddr, sOPCServerProgID, _wtof(sOPCServerConnTime));

	if(bOPCConnection)
	{
		m_StaticConnectStatus.SetWindowText(_T("Подключено"));
		m_BtnConnect.EnableWindow(FALSE);
		m_BtnDisconnect.EnableWindow(TRUE);

		AddItemToListBox(_T("Подключение - УСПЕХ"), 1);
	}
	else
	{AddItemToListBox(_T("Подключение - ОШИБКА"), 1);}

}



void CMFC_OPC_ExplorerDlg::OnBnClickedBtnDisconnect()
{
	// TODO: добавьте свой код обработчика уведомлений
/////////////////////////////////////////////
	if(bAddGroupTag)	RemoveGroup(pIOPCServer, hServerGroup);

/////////////////////////////////////////////
	AddItemToListBox(_T("Отключение..."), 1);
	bool bOPCDisConnection = DisConnect();

	if(bOPCDisConnection)
	{
		m_StaticConnectStatus.SetWindowText(_T("Отключено"));
		m_BtnDisconnect.EnableWindow(FALSE);
		m_BtnConnect.EnableWindow(TRUE);

		AddItemToListBox(_T("Отключение - УСПЕХ"), 1);
		bOPCConnection=false;
	}
	else
	{AddItemToListBox(_T("Отключение - ОШИБКА"), 1);}
	
}


void CMFC_OPC_ExplorerDlg::OnBnClickedBtnConnectRefresh()
{
	// TODO: добавьте свой код обработчика уведомлений
	//m_LstBxConnect.ResetContent();
	m_ListConnect.DeleteAllItems();
	//LstIndx.NULLIndxConnection();
}



////////////////////////////////////////////////////////////////
// 2 ///////////////////////////////////////////////////////////

void CMFC_OPC_ExplorerDlg::OnEnChangeEditGroupTagName()
{
	// TODO:  Добавьте код элемента управления
	m_EditGroupTagName.GetWindowText(sGroupTagName);
}


void CMFC_OPC_ExplorerDlg::OnBnClickedBtnGroupTagName()
{
	// TODO: добавьте свой код обработчика уведомлений

	if(bOPCConnection)
	{
		AddItemToListBox(_T("Добавление Группы..."), 2);
		bAddGroupTag = AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);

		if(bAddGroupTag)
		{
			AddItemToListBox(_T("Добавление Группы - УСПЕХ"), 2);
		}
		else
		{AddItemToListBox(_T("Добавление Группы - ОШИБКА"), 2);}
	}
	else
	{AddItemToListBox(_T("ОТСУСТВУЕТ СВЯЗЬ С ОРС-Сервером"), 2);}
}


////////////////////////////////////////////////////////////////

void CMFC_OPC_ExplorerDlg::OnEnChangeEditTagName()
{
	// TODO:  Добавьте код элемента управления
	m_EditTagName.GetWindowText(sTagName);
}


void CMFC_OPC_ExplorerDlg::OnBnClickedCheckTagType()
{
	// TODO: добавьте свой код обработчика уведомлений
	SetCmBxTagTypes();
}


void CMFC_OPC_ExplorerDlg::OnCbnSelchangeCmbxTagType()
{
	// TODO: добавьте свой код обработчика уведомлений
	//m_CmBxTagType.GetWindowText(TagType_CStr);
	int iCurItem = m_CmBxTagType.GetCurSel();
	m_CmBxTagType.GetLBText(iCurItem, sTagType);

}


void CMFC_OPC_ExplorerDlg::OnEnChangeEditTagType()
{
	// TODO:  Добавьте код элемента управления
	CString sEditTagType("");
	m_EditTagType.GetWindowText(sEditTagType);
	if( !sEditTagType.IsEmpty() )
	{
		sTagType=sEditTagType;
	}
}


void CMFC_OPC_ExplorerDlg::OnEnChangeEditTagValue()
{
	// TODO:  Добавьте код элемента управления
}


void CMFC_OPC_ExplorerDlg::OnBnClickedCheckTagReadWrite()
{
	// TODO: добавьте свой код обработчика уведомлений
	// обрабатываем ситуацию в сответсвие с проставленной галкой
    UINT nCheckReadWrite = m_CheckTagReadWrite.GetState();
    if(nCheckReadWrite &= BST_CHECKED)
    {
		// Писать
		m_BtnTagReadWrite.SetWindowText(_T("Писать [ V ]"));

		GetDlgItem(IDC_BTN_TAG_READ_CACHE)->EnableWindow(FALSE);
		GetDlgItem(IDC_BTN_TAG_READ_DEVICE)->EnableWindow(FALSE);
	}
    else //if(nCheck = BST_UNCHECKED)
    {
		// Читать
		m_BtnTagReadWrite.SetWindowText(_T("Читать [  ]"));

		GetDlgItem(IDC_BTN_TAG_READ_CACHE)->EnableWindow(TRUE);
		GetDlgItem(IDC_BTN_TAG_READ_DEVICE)->EnableWindow(TRUE);
	}

}


void CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadWrite()
{
	// TODO: добавьте свой код обработчика уведомлений
 	if(bOPCConnection)
	{
		if(nCheckReadWrite &= BST_CHECKED)
		{	// Писать
			AddItemToListBox(_T("Добавление Группы..."), 2);
			if(!bAddGroupTag)
			bAddGroupTag = AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);
			//WriteTag();

			if(bAddGroupTag)
			{
				AddItemToListBox(_T("Добавление Группы - УСПЕХ"), 2);
				AddItemToListBox(_T("Прописывание Тега..."), 2);
				bReadWriteTag = WriteTag();

				if(bReadWriteTag)	{AddItemToListBox(_T("Прописывание Тега - УСПЕХ"), 2);}
				else				{AddItemToListBox(_T("Прописывание Тега - ОШИБКА"), 2);}
			}
			else
			{AddItemToListBox(_T("Добавление Группы - ОШИБКА"), 2);}
		}
		else //if(nCheck = BST_UNCHECKED)
		{	// Читать
			AddItemToListBox(_T("Добавление Группы..."), 2);
			if(!bAddGroupTag)
			bAddGroupTag = AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);
			//ReadTag();

			if(bAddGroupTag)
			{
				AddItemToListBox(_T("Добавление Группы - УСПЕХ"), 2);
				AddItemToListBox(_T("Чтение Тега..."), 2);
				bReadWriteTag = ReadTag();

				if(bReadWriteTag)
				{
					AddItemToListBox(_T("Чтение Тега - УСПЕХ"), 2);

					CString sReadTagValue("");
					CString sReadTagQuality("");
					CString sReadTagProperties("");

					sReadTagValue = VariantValueToCStr(CurTag.TagVrntValue);
					sReadTagQuality = QualToFullCStr(CurTag.TagQuality);
					sReadTagProperties = OneTagPropsToFullCStr(pIOPCServer, CurTag);

					m_EditTagValue.SetWindowText(sReadTagValue);
					m_StaticTagQual.SetWindowText(sReadTagQuality);

					////////////////////////////////////////////////////////
					m_LstBxTagRWProperties.ResetContent();
					CString sPropertyPart(""), sNewCStr("\n");

					int dCStrLen = sReadTagProperties.GetLength();
					for(int i=0; i<dCStrLen; i++)
					{
						if(sReadTagProperties[i] != sNewCStr)
						{sPropertyPart += sReadTagProperties[i];}
						else
						{
							int nCount = m_LstBxTagRWProperties.GetCount();
							m_LstBxTagRWProperties.InsertString(nCount, sPropertyPart);
							sPropertyPart.Empty();
						}
					}

					int u=0;
				}
				else
				{AddItemToListBox(_T("Чтение Тега - ОШИБКА"), 2);}
			}
			else
			{AddItemToListBox(_T("Добавление Группы - ОШИБКА"), 2);}

/*
			AddItemToListBox(_T("Чтение Тега..."), 2);
			if(!bAddGroupTag)
			bAddGroupTag = AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);
			//ReadTag();

			if(bAddGroupTag)	{AddItemToListBox(_T("Чтение Тега - УСПЕХ"), 2);}
			else				{AddItemToListBox(_T("Чтение Тега - ОШИБКА"), 2);}
*/
		}
	}
	else
	{AddItemToListBox(_T("ОТСУСТВУЕТ СВЯЗЬ С ОРС-Сервером"), 2);}
}




#define OPC_QUALITY_DEVICE_TIMEOUT 0x0100

//SiteOverlayPictureType COnePicView::GetSiteSignalPicType( ItemValue& val )
OverlayPictureType CMFC_OPC_ExplorerDlg::GetSignalPicType( TagStruct& val )
{
	OverlayPictureType enPicType = OVL_EMPTY;

	if(V_VT(&val.TagVrntValue)==VT_BOOL)
	{
		switch(val.TagQuality&OPC_QUALITY_MASK)
		{
		case OPC_QUALITY_GOOD:
			enPicType = (val.TagQuality&OPC_QUALITY_DEVICE_TIMEOUT)? 
				(V_BOOL(&val.TagVrntValue)==VARIANT_TRUE? OVL_NONSIGNALED_TO : OVL_SIGNALED_TO)
				:
				(V_BOOL(&val.TagVrntValue)==VARIANT_TRUE? OVL_NONSIGNALED    : OVL_SIGNALED   );
			break;

		case OPC_QUALITY_BAD:
			switch(val.TagQuality&OPC_STATUS_MASK)
			{
			case OPC_QUALITY_LAST_KNOWN:
				enPicType = (val.TagQuality&OPC_QUALITY_DEVICE_TIMEOUT)? 
					(V_BOOL(&val.TagVrntValue)==VARIANT_TRUE? OVL_NONSIGNALED_TO : OVL_SIGNALED_TO)
					:
					(V_BOOL(&val.TagVrntValue)==VARIANT_TRUE? OVL_NONSIGNALED    : OVL_SIGNALED   );
				break;
			case OPC_QUALITY_SENSOR_FAILURE://"undef"
				enPicType = (val.TagQuality&OPC_QUALITY_DEVICE_TIMEOUT)? OVL_UNDEF_TO : OVL_UNDEF;
				break;
			case OPC_QUALITY_COMM_FAILURE://"error"
				enPicType = (val.TagQuality&OPC_QUALITY_DEVICE_TIMEOUT)? OVL_BREAK_TO : OVL_BREAK;
				break;
			default:
				enPicType = OVL_TIMEOUT;
			}
			break;
		}
	}

	return enPicType;
}
//*/



bool CMFC_OPC_ExplorerDlg::WriteTag()
{
	bool bRet(false);

//	#define TAGS_NUM 1
//	for(int i=0; i<TAGS_NUM; i++)
//	{
	// Add the OPC item
		int i=0;
		//std::vector<TagStruct> v_Tags(TAGS_NUM);
		//CurTag.TagName = sTagName;

		CurTag.TagFullName = sTagName;
		CurTag.TagVrntType = CStrVarTypeToVariantType(sTagType);
		//v_Tags[i][VCTR_INDX_INVRS]	 = sTagInvers;

		dAddItemTag = AddTheItem(pIOPCItemMgt, hServerItem, CurTag);

		if(dAddItemTag == 0)
		{
			//Prepare Item Value:
			CurTag.TagVrntValue = CStrToVariant(sTagValue, sTagType);
			//Write the value of the item to device:
			bWriteItemTag = WriteItem(pIOPCItemMgt, hServerItem, CurTag); 

			if(bWriteItemTag)
			{bRet=true;}
		}

		RemoveItem(pIOPCItemMgt, hServerItem);
//	}
	// Remove the OPC group: 
	// RemoveGroup(pIOPCServer, hServerGroup);

	return bRet;
}




bool CMFC_OPC_ExplorerDlg::ReadTag()
{
	bool bRet(false);

//	#define TAGS_NUM 1
//	for(int i=0; i<TAGS_NUM; i++)
//	{
	// Add the OPC item
		int i=0;
		//std::vector<TagStruct> v_Tags(TAGS_NUM);
		//CurTag.TagName = sTagName;

		CurTag.TagFullName = sTagName;
		CurTag.TagVrntType = CStrVarTypeToVariantType(sTagType);

		dAddItemTag = AddTheItem(pIOPCItemMgt, hServerItem, CurTag);

		if(dAddItemTag == 0)
		{
			//Read the value of the item from device:
			bReadItemTag = ReadItem(pIOPCItemMgt, hServerItem, CurTag); 

			if(bWriteItemTag)
			{bRet=true;}
		}

		RemoveItem(pIOPCItemMgt, hServerItem);

//	}
	// Remove the OPC group: 
	// RemoveGroup(pIOPCServer, hServerGroup);

	return bRet;
}



void CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadWriteRefresh()
{
	// TODO: добавьте свой код обработчика уведомлений
	m_ListTagReadWrite.DeleteAllItems();
	//LstIndx.NULLIndxRWTags();
}





////////////////////////////////////////////////////////////////
// 3 ///////////////////////////////////////////////////////////

void CMFC_OPC_ExplorerDlg::OnBnClickedBtnGetServerAddressSpace()
{
	// TODO: добавьте свой код обработчика уведомлений

//*
#if 0
	timl.Create(IDB_BRANCHES_LEAFS,16,0,RGB(255,255,255));
	m_TreeGetSrvrAddrSpace.SetImageList(&timl,TVSIL_NORMAL);

//	HTREEITEM hItem;
//	hItem = m_TreeGetSrvrAddrSpace.InsertItem(sBuf,hParent,TVI_LAST);

/*
	AddMenu(GetSubMenu(hMenu,i),hItem);
*/
/*
	TVINSERTSTRUCT tvInsert;
	tvInsert.hParent = NULL;
	tvInsert.hInsertAfter = NULL;
	//tvInsert.item.mask = TVIF_TEXT;
	tvInsert.item.mask = TVIF_TEXT;
	tvInsert.item.pszText = _T("United States");

	HTREEITEM hCountry = m_TreeGetSrvrAddrSpace.InsertItem(&tvInsert);
//*/


	TVITEM qitem;
	qitem.pszText = _T("United States");
//	qitem.hItem = hItem;
//	qitem.mask = TVIF_HANDLE|TVIF_PARAM|TVIF_IMAGE|TVIF_STATE|TVIF_SELECTEDIMAGE; // |TVIF_TEXT
	qitem.mask = TVIF_TEXT;
	qitem.stateMask = TVIS_STATEIMAGEMASK;
//	qitem.lParam = nID;
	qitem.iImage=0;
	qitem.state = INDEXTOSTATEIMAGEMASK(0);
	qitem.iSelectedImage = qitem.iImage;

	TVINSERTSTRUCT tvInsert;
	tvInsert.hParent = NULL;
	tvInsert.hInsertAfter = NULL;
	tvInsert.item = qitem;

	//m_TreeGetSrvrAddrSpace.SetItem(&item);
	HTREEITEM hCountry = m_TreeGetSrvrAddrSpace.InsertItem(&tvInsert); // +

	HTREEITEM hPA = m_TreeGetSrvrAddrSpace.InsertItem( // +
											// TVIF_TEXT
											//, _T("Pennsylvania")
											//, 0, 0, 0, 0, 0
											//, hCountry
											//, NULL

											//  TVIF_TEXT
											  _T("Pennsylvania")
											, 0, 0
											, hCountry
											, NULL
											);

	m_TreeGetSrvrAddrSpace.InsertItem(_T("Pittsburgh"), 1, 0, hPA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Harrisburg"), 1, 0, hPA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Altoona")	, 1, 0, hPA, TVI_SORT);	// -


	HTREEITEM hWA = m_TreeGetSrvrAddrSpace.InsertItem( // +
											//  TVIF_TEXT
											  _T("Washington")
											, 0, 0
											, hCountry
											, hPA
												);

	m_TreeGetSrvrAddrSpace.InsertItem(_T("Seattle"), 1, 0, hWA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Kalaloch"), 1, 0, hWA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Yakima"), 1, 0, hWA, TVI_SORT);		// -

#endif

	
/*
	timl.Create(IDB_BRANCHES_LEAFS,16,0,RGB(255,255,255));
	m_TreeGetSrvrAddrSpace.SetImageList(&timl,TVSIL_NORMAL);

	TVITEM qitem;
	qitem.pszText = _T("United States");
//	qitem.hItem = hItem;
//	qitem.mask = TVIF_HANDLE|TVIF_PARAM|TVIF_IMAGE|TVIF_STATE|TVIF_SELECTEDIMAGE; // |TVIF_TEXT
	qitem.mask = TVIF_TEXT;
	qitem.stateMask = TVIS_STATEIMAGEMASK;
//	qitem.lParam = nID;
	qitem.iImage=0;
	qitem.state = INDEXTOSTATEIMAGEMASK(0);
	qitem.iSelectedImage = qitem.iImage;

	TVINSERTSTRUCT tvInsert;
	tvInsert.hParent = NULL;
	tvInsert.hInsertAfter = NULL;
	tvInsert.item = qitem;

	//m_TreeGetSrvrAddrSpace.SetItem(&item);
	HTREEITEM hCountry = m_TreeGetSrvrAddrSpace.InsertItem(&tvInsert); // +

/*
	HTREEITEM hPA = m_TreeGetSrvrAddrSpace.InsertItem( // +
											// TVIF_TEXT
											//, _T("Pennsylvania")
											//, 0, 0, 0, 0, 0
											//, hCountry
											//, NULL

											//  TVIF_TEXT
											  _T("Pennsylvania")
											, 0, 0
											, hCountry
											, NULL
											);

	m_TreeGetSrvrAddrSpace.InsertItem(_T("Pittsburgh"), 1, 0, hPA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Harrisburg"), 1, 0, hPA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Altoona")	, 1, 0, hPA, TVI_SORT);	// -


	HTREEITEM hWA = m_TreeGetSrvrAddrSpace.InsertItem( // +
											//  TVIF_TEXT
											  _T("Washington")
											, 0, 0
											, hCountry
											, hPA
												);

	m_TreeGetSrvrAddrSpace.InsertItem(_T("Seattle"), 1, 0, hWA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Kalaloch"), 1, 0, hWA, TVI_SORT);	// -
	m_TreeGetSrvrAddrSpace.InsertItem(_T("Yakima"), 1, 0, hWA, TVI_SORT);		// -
//*/


////TryBrowseBranch(IBrowse, _T(""), _T(""));


 	if(bOPCConnection)
	{
		vv_OPCAddrSpace.clear();
		std::vector<CString> v_RootElems, v_OneElem;

		CComQIPtr<IOPCBrowseServerAddressSpace> IBrowse(pIOPCServer);
		if(IBrowse)
		{
//			TryBrowseBranch(IBrowse, _T(""), _T(""));

			CComPtr<IEnumString> IEnum;

			if(SUCCEEDED(
						hr=IBrowse->BrowseOPCItemIDs(
							OPC_BRANCH,//OPC_FLAT,
							L"",		//no vendor specific filtering
							VT_EMPTY,	//no data type filtering
							0,			//no Access Rights filtering
							&IEnum)
						)
				&& IEnum )
			{
				//"раскручиваем" перечислитель
				IEnum->Reset();

				LPOLESTR wsEnumString;
				while((hr = IEnum->Next(1,&wsEnumString,NULL))==S_OK)
				{
					//Есть строка!
					CString s("");
					s = wsEnumString;
					//ref_list.push_back(wsEnumString);
					
					v_RootElems.push_back(wsEnumString);
					v_OneElem.push_back(wsEnumString);
					vv_OPCAddrSpace.push_back(v_OneElem);
					v_OneElem.clear();

					//CleanUp
					::CoTaskMemFree(wsEnumString);
					int u(0);
				}

				IEnum.Release();
			}


			for(int i=0; i<v_RootElems.size(); i++)
			{
				if((hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_TO, v_RootElems[i]))==S_OK)
				{
					if(SUCCEEDED(
								hr=IBrowse->BrowseOPCItemIDs(
									OPC_LEAF,//OPC_FLAT,
									L"",		//no vendor specific filtering
									VT_EMPTY,	//no data type filtering
									0,			//no Access Rights filtering
									&IEnum)
								)
						&& IEnum )
					{
						//"раскручиваем" перечислитель
						IEnum->Reset();

						LPOLESTR wsEnumString;
						while((hr = IEnum->Next(1,&wsEnumString,NULL))==S_OK)
						{
							//Есть строка!
							CString s("");
							s = wsEnumString;
							//ref_list.push_back(wsEnumString);

							vv_OPCAddrSpace[i].push_back(wsEnumString);

							//CleanUp
							::CoTaskMemFree(wsEnumString);
							int u(0);
						}

						IEnum.Release();
					}
					else
					{
						//AppDiag.Record(2)<<"COPCServer::BrowseIOCItemIDs(OPC_LEAF) failed: "
						//	<<hres(hr)<<AppDiag.End();
						//return false;
					}
				}
			}

		IBrowse.Release();
		}

		v_RootElems.clear();
		v_OneElem.clear();


//////////////////////////////////////////////////////////////////////
//// После получения данных помещаем всё в элементы диалогового окна
//////////////////////////////////////////////////////////////////////
		timl.Create(IDB_BRANCHES_LEAFS, 16, 0, RGB(255,255,255));
		m_TreeGetSrvrAddrSpace.SetImageList(&timl,TVSIL_NORMAL);

	//*
		//std::vector<HTREEITEM> v_TreeRootElems;

		for(long i=0; i<vv_OPCAddrSpace.size(); i++)
		{
			TVITEM qitem;
			CString s = vv_OPCAddrSpace[i][0];

			USES_CONVERSION;		// Cstring to LPWSTR
			//LPWSTR psz = CT2W(s);
			LPWSTR psz = (LPWSTR)(LPCTSTR)s;

			qitem.pszText = psz; //_T("United States");
	//		qitem.hItem = hItem;
	//		qitem.mask = TVIF_HANDLE|TVIF_PARAM|TVIF_IMAGE|TVIF_STATE|TVIF_SELECTEDIMAGE; // |TVIF_TEXT
			qitem.mask = TVIF_TEXT;
			qitem.stateMask = TVIS_STATEIMAGEMASK;
	//		qitem.lParam = nID;
			qitem.iImage=0;
			qitem.state = INDEXTOSTATEIMAGEMASK(0);
			qitem.iSelectedImage = qitem.iImage;

			TVINSERTSTRUCT tvInsert;
			tvInsert.hParent = NULL;
			tvInsert.hInsertAfter = NULL;
			tvInsert.item = qitem;

			//m_TreeGetSrvrAddrSpace.SetItem(&item);
			//v_TreeRootElems.push_back( m_TreeGetSrvrAddrSpace.InsertItem(&tvInsert) ); // +
			HTREEITEM htrCurTreeRootElem = m_TreeGetSrvrAddrSpace.InsertItem(&tvInsert);

			for(long j=1; j<vv_OPCAddrSpace[i].size(); j++) // Начинаем с 
			{
			HTREEITEM hPA = m_TreeGetSrvrAddrSpace.InsertItem( // +
														//  TVIF_TEXT
														  vv_OPCAddrSpace[i][j] //_T("Pennsylvania")
														, 1, 0
														, htrCurTreeRootElem
														, TVI_SORT
														);
			}

		}


	/*
			HTREEITEM hPA = m_TreeGetSrvrAddrSpace.InsertItem( // +
													// TVIF_TEXT
													//, _T("Pennsylvania")
													//, 0, 0, 0, 0, 0
													//, hCountry
													//, NULL

													//  TVIF_TEXT
													  _T("Pennsylvania")
													, 0, 0
													, hCountry
													, NULL
													);

			m_TreeGetSrvrAddrSpace.InsertItem(_T("Pittsburgh"), 1, 0, hPA, TVI_SORT);	// -
			m_TreeGetSrvrAddrSpace.InsertItem(_T("Harrisburg"), 1, 0, hPA, TVI_SORT);	// -
			m_TreeGetSrvrAddrSpace.InsertItem(_T("Altoona")	, 1, 0, hPA, TVI_SORT);	// -


			HTREEITEM hWA = m_TreeGetSrvrAddrSpace.InsertItem( // +
													//  TVIF_TEXT
													  _T("Washington")
													, 0, 0
													, hCountry
													, hPA
														);

			m_TreeGetSrvrAddrSpace.InsertItem(_T("Seattle"), 1, 0, hWA, TVI_SORT);	// -
			m_TreeGetSrvrAddrSpace.InsertItem(_T("Kalaloch"), 1, 0, hWA, TVI_SORT);	// -
			m_TreeGetSrvrAddrSpace.InsertItem(_T("Yakima"), 1, 0, hWA, TVI_SORT);		// -
	//*/


#if 0
	/////////////////////////////////////////////////////////////////
		//Адресное пространство сервера
		//запросить интерфейс
		CComQIPtr<IOPCBrowseServerAddressSpace> IBrowse(pIOPCServer);
		if(IBrowse)
		{
			OPCNAMESPACETYPE onst;
			IBrowse->QueryOrganization(&onst);

			CComPtr<IEnumString> IEnum;
			CComPtr<IEnumString> IIEnum;

			hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_UP, _T(""));
			hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_DOWN, _T(""));


//			while( (hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_UP, _T(""))) == S_OK )
//			while( (hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_DOWN, _T(""))) == S_OK )
//			{
//			}


			if(SUCCEEDED(
						hr=IBrowse->BrowseOPCItemIDs(
							OPC_BRANCH
							//OPC_LEAF
							//OPC_FLAT
							, L""		//no vendor specific filtering
							, VT_EMPTY	//no data type filtering
							, 0			//no Access Rights filtering
							, &IEnum)
						)
				&& IEnum )
			{
				//"раскручиваем" перечислитель
				IEnum->Reset();

				hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_UP, _T(""));
				hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_DOWN, _T(""));


				hr=IBrowse->BrowseOPCItemIDs(
							OPC_LEAF
							//OPC_FLAT
							, L""		//no vendor specific filtering
							, VT_EMPTY	//no data type filtering
							, 0			//no Access Rights filtering
							, &IEnum
							);


				LPOLESTR wsEnumString;
				while((hr = IEnum->Next(1,&wsEnumString,NULL))==S_OK)
				{
					//Есть строка!
					CString s("");
					s = wsEnumString;

//TryBrowseOneBranch(IBrowse, _T(""), _T(""));

					v_RootElems.push_back(s);



						hr=IBrowse->BrowseOPCItemIDs(
							OPC_BRANCH, //OPC_BRANCH,//OPC_FLAT, // OPC_LEAF
							L"",		//no vendor specific filtering
							VT_EMPTY,	//no data type filtering
							0,			//no Access Rights filtering
							&IEnum);



					IEnum->Reset();
					while((hr = IEnum->Next(1,&wsEnumString,NULL))==S_OK)
					{
						CString s("");
						s = wsEnumString;
					}
//////////////////////////////////////////////////////////////////////////////
/* // Получение плоской иерархии всего тегового пространства. 11-111-1112-21-22-23-24-241-242-31-32-321-322-33

					IOPCBrowseServerAddressSpace *pBSAS = IBrowse;
					LPCWSTR lpprestr;
					LPCWSTR lpcw=_T("");
					HRESULT r1(NULL);
					WCHAR wchstr;
					IEnumString *pEEnum = 0;

					HRESULT r2(NULL);
					LPOLESTR pStr;
					ULONG actual;

					pBSAS->ChangeBrowsePosition(OPC_BROWSE_DOWN, lpcw);
					r1 = pBSAS->BrowseOPCItemIDs(OPC_FLAT, &wchstr, VT_EMPTY, 0, &pEEnum); // OPC_BRANCH, OPC_LEAF, OPC_FLAT

					while( (r1==S_OK) && (r2 = pEEnum->Next(1, &pStr, &actual)) == S_OK)
					{
						CString s("");
						//s.Format(_T("%ls-%ls--%ls"), lpprestr, lpcw, pStr);
						s.Format(_T("%ls"), pStr);
						v_RootElems.push_back(s);
					}

					pEEnum->Release();
					pBSAS->ChangeBrowsePosition(OPC_BROWSE_UP, lpcw);
//*/


//					IOPCBrowseServerAddressSpace *pBSAS = IBrowse;

//					IBrowse->ChangeBrowsePosition(OPC_BROWSE_DOWN, _T(""));

//					pBSAS->ChangeBrowsePosition(OPC_BROWSE_UP, _T(""));
/*
					LPOLESTR wsEEnumString;
					ULONG actual;
					pBSAS->BrowseAccessPaths(L"", &IIEnum);
					while((hr = IEnum->Next(1, &wsEEnumString, &actual)) == S_OK)
					{
						CString s;
						s = wsEEnumString;
						//sss.Format(_T("%ls-%ls--%ls---%ls"), lpprestr, lp, lpcw, pStr);
						int u(0);
					}
*/

//*
//					while( (hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_DOWN, _T(""))) == S_OK )
					while( (hr = IBrowse->ChangeBrowsePosition(OPC_BROWSE_UP, _T(""))) == S_OK )
					{

					}

					if(SUCCEEDED(
								hr=IBrowse->BrowseOPCItemIDs(
									OPC_LEAF,//OPC_FLAT,
									L"",		//no vendor specific filtering
									VT_EMPTY,	//no data type filtering
									0,			//no Access Rights filtering
									&IIEnum)
								)
						&& IIEnum )
					{
						LPOLESTR wsEEnumString;
						ULONG actual;
//						while((r2 = IIEnum->Next(1, &wsEEnumString, &actual)) == S_OK)
						while((hr = IIEnum->Next(1, &wsEEnumString, &actual)) == S_OK)
						{
							CString s;
							s = wsEEnumString;
							//sss.Format(_T("%ls-%ls--%ls---%ls"), lpprestr, lp, lpcw, pStr);
							int u(0);
						}
							int u(0);

						//"раскручиваем" перечислитель
						
						//IIEnum->Reset();
						//
						//LPOLESTR wsEEnumString;
						//while((hr = IIEnum->Next(1,&wsEEnumString,NULL))==S_OK)
						//{
						//	//Есть строка!
						//	CString s("");
						//	s = wsEEnumString;
						//	//ref_list.push_back(wsEEnumString);
						//	v_RootElems.push_back(s);
						//
						//	//CleanUp
						//	::CoTaskMemFree(wsEEnumString);
						//}
						
					}
					else
					{
						//AppDiag.Record(2)<<"COPCServer::BrowseIOCItemIDs(OPC_LEAF) failed: "
						//	<<hres(hr)<<AppDiag.End();
						//return false;
					}
					//ref_list.push_back(wsEEnumString);

//*/
//////////////////////////////////////////////////////////////////////////////

					if(v_RootElems.size()!=0)
					{
						vv_OPCAddrSpace.push_back(v_RootElems);
						v_RootElems.clear();
					}

					//CleanUp
					::CoTaskMemFree(wsEnumString);
				}
			}
			else
			{
				//AppDiag.Record(2)<<"COPCServer::BrowseIOCItemIDs(OPC_LEAF) failed: "
				//	<<hres(hr)<<AppDiag.End();
				//return false;
			}
		}
		else
		{
			//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCBrowseServerAddressSpace): "
			//		<<hres(hr)<<AppDiag.End();
			//return false;
		}

		int u(0);
//*


//	pBSAS->ChangeBrowsePosition(OPC_BROWSE_UP, lpcw);
//	pBSAS->ChangeBrowsePosition(OPC_BROWSE_DOWN, lpcw);
//	r1 = pBSAS->BrowseOPCItemIDs(OPC_FLAT, &s, VT_EMPTY, 0, &pEnum);
#endif



//*/

#if 0

		// иерархия Item-ов сервера деревом/простой список (OPC_NS_HIERARCHIAL / OPC_NS_FLAT) это можно узнать вызвав QueryOrganization
		// ChangeBrowsePosition, BrowseOPCItemIDs, GetItemID
		// 2а) если просто список: то ChangeBrowsePosition тебе никак не нужно а получить энумератор всех Item-ов (IEnumString) можно через BrowseOPCItemIDs. А дальше дерево можешь строить (или не строить) сам.
		// 2б) если дерево: я бы тебе предложил простой способ охода: поднимаешься на самый верх дерева (ChangeBrowsePosition(OPC_BROWSE_UP, _T("")) до упора) 
		// 3) потом получаешь через BrowseOPCItemIDs все листья (OPC_LEAF) и дочерние узлы (OPC_BRANCH) — добавляешь у себя в дерево. 
		// 4) потом повторяешь для каждого дочеренего узла операцию 3 пока не остануться одни листья. 
		// P.S. GetItemID — можно вызывать для листьев что бы точно узнать название Item-а

	//	OPC_BRANCH;
	//	hr=IBrowse->BrowseOPCItemIDs( OPC_BRANCH, ...);

		//instance OpcServer;
		//value OPCNAMESPACETYPE;
		OPCNAMESPACETYPE value;

	//	value = pIOPCServer->QueryInterface();

	/////////////////////////////////////////////////////////////////

	CComQIPtr<IOPCBrowseServerAddressSpace> IBrowse(pIOPCServer);
	if(IBrowse)
	{
	TryBrowseBranch(IBrowse, _T(""), _T(""));
	}

	/////////////////////////////////////////////////////////////////
		//Адресное пространство сервера
		//запросить интерфейс
		CComQIPtr<IOPCBrowseServerAddressSpace> IBrowse(pIOPCServer);
		if(IBrowse)
		{
			TryBrowseBranch(IBrowse, _T(""), _T(""));

			CComPtr<IEnumString> IEnum;

			if(SUCCEEDED(
						hr=IBrowse->BrowseOPCItemIDs(
							OPC_BRANCH,//OPC_FLAT,
							L"",		//no vendor specific filtering
							VT_EMPTY,	//no data type filtering
							0,			//no Access Rights filtering
							&IEnum)
						)
				&& IEnum )
			{
				//"раскручиваем" перечислитель
				IEnum->Reset();

				LPOLESTR wsEnumString;
				while((hr = IEnum->Next(1,&wsEnumString,NULL))==S_OK)
				{
					//Есть строка!
					CString s("");
					s = wsEnumString;
					//ref_list.push_back(wsEnumString);

					//CleanUp
					::CoTaskMemFree(wsEnumString);
					int u(0);
				}
			}


			if(SUCCEEDED(
						hr=IBrowse->BrowseOPCItemIDs(
							OPC_LEAF,//OPC_FLAT,
							L"",		//no vendor specific filtering
							VT_EMPTY,	//no data type filtering
							0,			//no Access Rights filtering
							&IEnum)
						)
				&& IEnum )
			{
				//"раскручиваем" перечислитель
				IEnum->Reset();

				LPOLESTR wsEnumString;
				while((hr = IEnum->Next(1,&wsEnumString,NULL))==S_OK)
				{
					//Есть строка!
					CString s("");
					s = wsEnumString;
					//ref_list.push_back(wsEnumString);

					//CleanUp
					::CoTaskMemFree(wsEnumString);
					int u(0);
				}
			}
			else
			{
				//AppDiag.Record(2)<<"COPCServer::BrowseIOCItemIDs(OPC_LEAF) failed: "
				//	<<hres(hr)<<AppDiag.End();
				//return false;
			}
		}
		else
		{
			//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCBrowseServerAddressSpace): "
			//		<<hres(hr)<<AppDiag.End();
			//return false;
		}
#endif

	}
	else
	{AddItemToListBox(_T("ОТСУСТВУЕТ СВЯЗЬ С ОРС-Сервером"), 3);}

}



void CMFC_OPC_ExplorerDlg::TryBrowseOneBranch(IOPCBrowseServerAddressSpace *pBSAS, LPCWSTR lpprestr, LPCWSTR lpcw)
{
	HRESULT r1(NULL);
	//OPCNAMESPACETYPE onst;
	//LPWSTR pItemID;
	WCHAR wchstr;
	//WCHAR *lp;
	IEnumString *pEnum = 0;

	HRESULT r2(NULL);
	//char szBuf[40];
	//long nPos = 1;
	LPOLESTR pStr;
	ULONG actual;

	pBSAS->ChangeBrowsePosition(OPC_BROWSE_DOWN, lpcw);
	r1 = pBSAS->BrowseOPCItemIDs(OPC_FLAT, &wchstr, VT_EMPTY, 0, &pEnum);

	while( (r1==S_OK) && (r2 = pEnum->Next(1, &pStr, &actual)) == S_OK)
	{
		CString s("");
		//s.Format(_T("%ls-%ls--%ls"), lpprestr, lpcw, pStr);
		s.Format(_T("%ls"), pStr);
	}

	pEnum->Release();
	pBSAS->ChangeBrowsePosition(OPC_BROWSE_UP, lpcw);
}



void CMFC_OPC_ExplorerDlg::TryBrowseBranch(IOPCBrowseServerAddressSpace *pBSAS, LPCWSTR lpprestr, LPCWSTR lpcw)
{
//	IOPCBrowseServerAddressSpace *pBSAS;
//	LPCWSTR lpprestr;
//	LPCWSTR lpcw;

	HRESULT r1;
	OPCNAMESPACETYPE onst;
	LPWSTR pItemID;
	WCHAR s, *lp;
	IEnumString *pEnum = 0;

	HRESULT r2 = S_OK;
	char szBuf[40];
	long nPos = 1;
	LPOLESTR pStr;
	ULONG actual;

	//lp = WSTRFromSBC("%ls,%ls", pIMalloc);
	//pBSAS->ChangeBrowsePosition(OPC_BROWSE_UP, lpcw);

	pBSAS->ChangeBrowsePosition(OPC_BROWSE_DOWN, lpcw);
	r1 = pBSAS->BrowseOPCItemIDs(OPC_FLAT, &s, VT_EMPTY, 0, &pEnum);
	pBSAS->QueryOrganization(&onst);

	switch(onst)
	{
	case OPC_NS_HIERARCHIAL:
		while((r2 = pEnum->Next(1, &pStr, &actual)) == S_OK)
		{
			CString sss;
			sss.Format(_T("%ls-%ls--%ls"), lpprestr, lpcw, pStr);
			//sss.Format(_T("%ls-%ls--%ls---%ls"), lpprestr, lp, lpcw, pStr);

		//	if(onst == OPC_BRANCH)
		//	TryBrowsBeranch(pBSAS, lpprestr, lpcw);
			
		}
		break;

	case OPC_NS_FLAT:
		while((r2 = pEnum->Next(1, &pStr, &actual)) == S_OK)
		{
			CString sss;
			sss.Format(_T("%ls-%ls--%ls"), lpprestr, lpcw, pStr);
			//sss.Format(_T("%ls-%ls--%ls---%ls"), lpprestr, lp, lpcw, pStr);

			if(onst == OPC_BRANCH)
			TryBrowseBranch(pBSAS, lpprestr, lpcw);
			
		}
		break;
	}

/*
	while((r2 = pEnum->Next(1, &pStr, &actual)) == S_OK)
	{
		CString sss;
		sss.Format(_T("%ls-%ls--%ls"), lpprestr, lpcw, pStr);
		//sss.Format(_T("%ls-%ls--%ls---%ls"), lpprestr, lp, lpcw, pStr);

		if(onst == OPC_BRANCH)
		TryBrowseBranch(pBSAS, lpprestr, lpcw);
			
	}
*/

	//pBSAS->ChangeBrowsePosition(OPC_BROWSE_DOWN, lpcw);
	// OPC_NS_HIERARCHIAL / OPC_NS_FLAT

	// get first item

	//pIMalloc->Free(lp);
	//pIMalloc->Free(pStr);
	pEnum->Release();

	pBSAS->ChangeBrowsePosition(OPC_BROWSE_UP, lpcw);
}


void CMFC_OPC_ExplorerDlg::OnTvnSelchangedTreeServerAddressSpace(NMHDR *pNMHDR, LRESULT *pResult)
{
/////	LPNMTVKEYDOWN pTVKeyDown = reinterpret_cast<LPNMTVKEYDOWN>(pNMHDR);
/////	*pResult = 0;
//	m_LstBxTagProperties;

	map_TagProps.clear();

//	LPNMTVKEYDOWN pTVKeyDown = reinterpret_cast<LPNMTVKEYDOWN>(pNMHDR);
	HTREEITEM hItem = m_TreeGetSrvrAddrSpace.GetSelectedItem();
	if(hItem!=NULL)
	{
		HTREEITEM htrRootItem = m_TreeGetSrvrAddrSpace.GetRootItem();
		HTREEITEM htrCrntItem = m_TreeGetSrvrAddrSpace.GetSelectedItem();
		HTREEITEM htrPrntItem = m_TreeGetSrvrAddrSpace.GetParentItem(htrCrntItem);

		CString  sCurItemText = m_TreeGetSrvrAddrSpace.GetItemText(htrCrntItem);
		CString  sFullTagName(sCurItemText);

		//while( (htrRootItem != htrPrntItem) && (htrPrntItem) )
		//while( (htrRootItem != htrCrntItem) && (htrPrntItem) )
		while( htrPrntItem )
		{
			htrCrntItem = htrPrntItem;
			htrPrntItem = m_TreeGetSrvrAddrSpace.GetParentItem(htrCrntItem);
			sCurItemText = m_TreeGetSrvrAddrSpace.GetItemText(htrCrntItem);
			sFullTagName.Insert(0, _T("/"));
			sFullTagName.Insert(0, sCurItemText);
			sCurItemText.Empty();
		}


		CString  text = sFullTagName;


 		if(bOPCConnection)
		{
			try{
				HRESULT hr;

				//Адресное пространство сервера
				//запросить интерфейс
				IOPCItemProperties* pIProps = NULL;
				hr = pIOPCServer->QueryInterface(IID_IOPCItemProperties,(void**)&pIProps);
				if(FAILED(hr)) {
					//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCItemProperties): "
					//		<<hres(hr)<<AppDiag.End();
					throw hr;
				}

				//запросить список значений свойств тэга
				# define PROPS_NUM 6

				DWORD adwProps[PROPS_NUM] = { 100, 101, 5000, 5001, 5002, 5003 };
				/*DWORD adwProps[11] = {
				1 - Canonical Data Type	- Type
				2 - Item Value			- Val
				3 - Item Quality		- Short Interger
				4 - Item Timestamp		- Date
				5 - Item Access Rights	- Long Integer
				100 - EU Units			- String
				101 - Item Description	- String
				5000 - Associated Product Name - String
				5001 - Physical quantity - String
				5002 - Site Description	- String
				5003 - Site ID			- Unsigned Long
				};
				*/

				VARIANT* pPropValues = NULL;
				HRESULT* pPropValuesErrors = NULL;

				USES_CONVERSION;		// Cstring to LPWSTR
				//LPWSTR psz = CT2W(s);
				LPWSTR psz = (LPWSTR)(LPCTSTR)text;

				hr = pIProps->GetItemProperties(
												//(LPWSTR)strTag.c_str()
												psz
												, PROPS_NUM
												, adwProps
												, &pPropValues
												, &pPropValuesErrors
											  );
				if(FAILED(hr)) {
					//AppDiag.Record(2)<<"COPCServer::GetItemProperties failed: "
					//		<<hres(hr)<<AppDiag.End();
					pIProps->Release();
					throw hr;
				}


				for(DWORD dwI=0; dwI<PROPS_NUM; dwI++)
				{
					if(pPropValuesErrors[dwI]==S_OK)
					{
						//Props[ adwProps[dwI] ] = pPropValues[dwI];
						//v_Props.push_back( VariantToCStr(pPropValues[dwI]) );
						map_TagProps[adwProps[dwI]] = ( VariantValueToCStr(pPropValues[dwI]) );
					}
				}

				//освобождаем выделенную сервером память
				::CoTaskMemFree(pPropValues);
				::CoTaskMemFree(pPropValuesErrors);

				//освобождаем интерфейс
				pIProps->Release();
			}
			catch(HRESULT hError)
			{
				hError;

//				return false;
			}

//			return true;
		}
		else
		{AddItemToListBox(_T("ОТСУСТВУЕТ СВЯЗЬ С ОРС-Сервером"), 3);}


//////////////////////////////////////////////////////////////////////
//// После получения данных помещаем всё в элементы диалогового окна
//////////////////////////////////////////////////////////////////////
		m_LstBxTagProperties.ResetContent();

		for(TagPropsIter = map_TagProps.begin(); TagPropsIter != map_TagProps.end(); TagPropsIter++)
		{
			// к ключу можно обращаться еще вот так  //(*iter).first и (*iter).second соответственно
			CString sTagProperties("");
			sTagProperties = sTagProperties
				+ LUintToCStr(TagPropsIter->first)
				+ _T(": ")
				+ TagPropsIter->second
				;

			int nCount = m_LstBxTagProperties.GetCount();
			m_LstBxTagProperties.InsertString(nCount, sTagProperties);
		}

		// handle state change here or post message to another handler
		// Post message state has changed…
		//PostMessage(TREE_VIEW_CHECK_STATE_CHANGE,0,(LPARAM)item);
		int u=0;

	}


	*pResult = 0;

/*
/////////////////////////////////////////////////////////////
//	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	LPNMTVKEYDOWN pTVKeyDown = reinterpret_cast<LPNMTVKEYDOWN>(pNMHDR);
	// TODO: Add your control notification handler code here
	// we only want to handle the space key press
	if(pTVKeyDown->wVKey==VK_SPACE)
	{
		// Determine the selected tree item.
		HTREEITEM item = m_TreeGetSrvrAddrSpace.GetSelectedItem();
		if(item!=NULL)
		{
			item;
			// handle state change here or post message to another handler
			// Post message state has changed…
			//PostMessage(TREE_VIEW_CHECK_STATE_CHANGE,0,(LPARAM)item);
		}
	}
	*pResult = 0;

/////////////////////////////////////////////////////////////
	HTREEITEM changedItem(NULL);
	UINT nID = (UINT)m_TreeGetSrvrAddrSpace.GetItemData(changedItem);
*/

}


void CMFC_OPC_ExplorerDlg::OnBnClickedBtnServerAddressSpaceRefresh()
{
	// TODO: добавьте свой код обработчика уведомлений
	m_ListGetSrvrAddrSpace.DeleteAllItems();
	//LstIndx.NULLIndxSrvrAddrSp();
}


///////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////

void CMFC_OPC_ExplorerDlg::AddItemToListBox(CString sStrToAdd, int nNumOfLstBx)
{
	int curIndex=0;	// Индекс элемента выбранного в box
	int nCount=0;	// Число элементов внутри списка
	
	switch(nNumOfLstBx)
	{
	case 1:
		nCount =  m_ListConnect.GetItemCount();
		m_ListConnect.InsertItem(nCount, LUintToCStr(nCount) +_T(" - ")+ sStrToAdd);
		break;
	case 2:
		nCount =  m_ListTagReadWrite.GetItemCount();
		m_ListTagReadWrite.InsertItem(nCount, LUintToCStr(nCount) +_T(" - ")+ sStrToAdd);
		break;
	case 3:
		nCount =  m_ListGetSrvrAddrSpace.GetItemCount();
		m_ListGetSrvrAddrSpace.InsertItem(nCount, LUintToCStr(nCount) +_T(" - ")+ sStrToAdd);
		break;
	}
}



///////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////


bool CMFC_OPC_ExplorerDlg::GetServerTags( std::list<std::wstring>& ref_list )
{
	HRESULT hr = S_OK;

	ref_list.clear();

#if 1
	//Адресное пространство сервера
	//запросить интерфейс
	CComQIPtr<IOPCBrowseServerAddressSpace> IBrowse(pIOPCServer);
	if(IBrowse)
	{
		CComPtr<IEnumString> IEnum;
		if(SUCCEEDED(
					hr=IBrowse->BrowseOPCItemIDs(
						OPC_LEAF,//OPC_FLAT,
						L"",		//no vendor specific filtering
						VT_EMPTY,	//no data type filtering
						0,			//no Access Rights filtering
						&IEnum)
					)
			&& IEnum )
		{
			//"раскручиваем" перечислитель
			IEnum->Reset();

			LPOLESTR wsEnumString;
			while((hr = IEnum->Next(1,&wsEnumString,NULL))==S_OK)
			{
				//Есть строка!
				ref_list.push_back(wsEnumString);

				//CleanUp
				::CoTaskMemFree(wsEnumString);
			}
		}
		else
		{
			//AppDiag.Record(2)<<"COPCServer::BrowseIOCItemIDs(OPC_LEAF) failed: "
			//	<<hres(hr)<<AppDiag.End();
			return false;
		}
	}
	else
	{
		//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCBrowseServerAddressSpace): "
		//		<<hres(hr)<<AppDiag.End();
		return false;
	}

	//умные указатели сделают свое дело при выходе из области видимости
#else
	try{

		//Адресное пространство сервера
		//запросить интерфейс
		IOPCBrowseServerAddressSpace* pIBrowse = NULL;
		hr = m_pOPCServerInterface->QueryInterface(IID_IOPCBrowseServerAddressSpace,(void**)&pIBrowse);
		if(FAILED(hr)) {
			AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCBrowseServerAddressSpace): "
					<<hres(hr)<<AppDiag.End();
			throw hr;
		}

		//запросить адресное пространство в виде узлов корня
		IEnumString* pIEnum = NULL;
		hr = pIBrowse->BrowseOPCItemIDs(
						OPC_LEAF,//OPC_FLAT,
						L"",		//no vendor specific filtering
						VT_EMPTY,	//no data type filtering
						0,			//no Access Rights filtering
						&pIEnum);
		if(FAILED(hr)) {
			AppDiag.Record(2)<<"COPCServer::BrowseIOCItemIDs(OPC_LEAF) failed: "
					<<hres(hr)<<AppDiag.End();
			pIBrowse->Release();
			throw hr;
		}

		//"раскручиваем" перечислитель
		pIEnum->Reset();

		LPOLESTR wsEnumString;
		while((hr = pIEnum->Next(1,&wsEnumString,NULL))==S_OK)
		{
			//Есть строка!
			ref_list.push_back(wsEnumString);

			//CleanUp
			::CoTaskMemFree(wsEnumString);
		}
			
		//мавр сделал свое дело - мавр может уходить...
		pIEnum->Release();

		pIBrowse->Release();
	}
	catch(HRESULT hError)
	{
		hError;

		return false;
	}
#endif

	return ref_list.size()>0;
}



bool CMFC_OPC_ExplorerDlg::GetSites( std::list<std::wstring>& ref_list )
{
	ref_list.clear();

	try{
		HRESULT hr;

		//Адресное пространство сервера
		//запросить интерфейс

		IOPCBrowseServerAddressSpace* pIBrowse = NULL;
		hr = pIOPCServer->QueryInterface(IID_IOPCBrowseServerAddressSpace,(void**)&pIBrowse);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCBrowseServerAddressSpace): "
			//		<<hres(hr)<<AppDiag.End();
			throw hr;
		}

		IOPCItemProperties* pIProps = NULL;
		hr = pIOPCServer->QueryInterface(IID_IOPCItemProperties,(void**)&pIProps);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCItemProperties): "
			//		<<hres(hr)<<AppDiag.End();
			throw hr;
		}

/*
		//перейти к ветви в дереве адресного пространства
		hr = pIBrowse->ChangeBrowsePosition(OPC_BROWSE_TO, L"Sites");
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer::ChangeBrowsePosition(OPC_BROWSE_TO, \"Sites\") failed: "
			//		<<hres(hr)<<AppDiag.End();
			pIBrowse->Release();
			throw hr;
		}
*/
		//запросить адресное пространство в виде подветвей выбранной ветви
		IEnumString* pIEnum = NULL;
		hr = pIBrowse->BrowseOPCItemIDs(
						OPC_BRANCH,
						L"",		//no vendor specific filtering
						VT_EMPTY,	//no data type filtering
						0,			//no Access Rights filtering
						&pIEnum);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer::BrowseIOCItemIDs(OPC_BRANCH) failed: "
			//		<<hres(hr)<<AppDiag.End();
			pIBrowse->Release();
			throw hr;
		}

		//"раскручиваем" перечислитель
		LPOLESTR wsEnumString;

		pIEnum->Reset();
		while((hr = pIEnum->Next(1,&wsEnumString,NULL))==S_OK)
		{
			//Есть строка!
			ref_list.push_back(wsEnumString);

			std::wstring strSiteID;
			strSiteID += L"Sites/";
			strSiteID += wsEnumString;
			strSiteID += L"/ID";

			//CleanUp
			::CoTaskMemFree(wsEnumString);

			VARIANT* pPropValues = NULL;
			HRESULT* pPropValuesErrors = NULL;
			DWORD dwValuePropID = 2;

			/*
			hr = pIProps->GetItemProperties((LPWSTR)strSiteID.c_str(),
											1, &dwValuePropID,
											&pPropValues,
											&pPropValuesErrors);
			if(FAILED(hr)) {
			//	AppDiag.Record(2)<<"COPCServer::GetItemProperties failed: "
			//			<<hres(hr)<<AppDiag.End();
				pIEnum->Release();
				pIBrowse->Release();
				pIProps->Release();
				throw hr;
			}
			*/

			//UINT uSiteID = V_UI4(&pPropValues[0]);

			//освобождаем выделенную сервером память
			::CoTaskMemFree(pPropValues);
			::CoTaskMemFree(pPropValuesErrors);
		}
			
		//освобождаем интерфейсы
		pIEnum->Release();
		pIBrowse->Release();
		pIProps->Release();

		//////////////////////////////////////////////////////////////////////////
	}
	catch(HRESULT hError)
	{
		hError;

		return false;
	}

	return ref_list.size()>0;
}



bool CMFC_OPC_ExplorerDlg::GetTagProperties( const std::wstring& strTag, mapTagProps& Props )
{
	try{
		HRESULT hr;

		//Адресное пространство сервера
		//запросить интерфейс
		IOPCItemProperties* pIProps = NULL;
		hr = pIOPCServer->QueryInterface(IID_IOPCItemProperties,(void**)&pIProps);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCItemProperties): "
			//		<<hres(hr)<<AppDiag.End();
			throw hr;
		}

		//запросить список значений свойств тэга
		DWORD adwProps[6] = { 100, 101, 5000, 5001, 5002, 5003 };

		VARIANT* pPropValues = NULL;
		HRESULT* pPropValuesErrors = NULL;

		hr = pIProps->GetItemProperties((LPWSTR)strTag.c_str(),
										6, adwProps,
										&pPropValues,
										&pPropValuesErrors);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer::GetItemProperties failed: "
			//		<<hres(hr)<<AppDiag.End();
			pIProps->Release();
			throw hr;
		}

		for(DWORD dwI=0; dwI<6; dwI++)
		{
			if(pPropValuesErrors[dwI]==S_OK)
			{
				Props[ adwProps[dwI] ] = pPropValues[dwI];
			}
		}

		//освобождаем выделенную сервером память
		::CoTaskMemFree(pPropValues);
		::CoTaskMemFree(pPropValuesErrors);

		//освобождаем интерфейс
		pIProps->Release();
	}
	catch(HRESULT hError)
	{
		hError;

		return false;
	}

	return true;
}

bool CMFC_OPC_ExplorerDlg::GetTagProperty( const std::wstring& strTag, DWORD dwPropID, CComVariant& varProp )
{
	try{
		HRESULT hr;

		//запросить интерфейс работы со свойствами
		IOPCItemProperties* pIProps = NULL;
		hr = pIOPCServer->QueryInterface(IID_IOPCItemProperties,(void**)&pIProps);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCItemProperties): "
			//		<<hres(hr)<<end;
			throw hr;
		}

		//запросить свойство тэга
		VARIANT* pPropValues = NULL;
		HRESULT* pPropValuesErrors = NULL;

		hr = pIProps->GetItemProperties((LPWSTR)strTag.c_str(),
										1, &dwPropID,
										&pPropValues,
										&pPropValuesErrors);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer::GetItemProperties failed: "
			//		<<hres(hr)<<AppDiag.End();
			pIProps->Release();
			throw hr;
		}

		if(pPropValuesErrors[0]==S_OK)
		{
			varProp = pPropValues[0];
		}

		//освобождаем выделенную сервером память
		::CoTaskMemFree(pPropValues);
		::CoTaskMemFree(pPropValuesErrors);

		//освобождаем интерфейс
		pIProps->Release();
	}
	catch(HRESULT hError)
	{
		hError;
		return false;
	}

	return true;
}


///////////////////////////////////////////////////////////////////////////////////////
// УПАКОВАННЫЕ МЕТОДЫ РЕАЛИЗАЦИЯ И Т.Д. ///////////////////////////////////////////////

bool CMFC_OPC_ExplorerDlg::Connect(
				  LPCTSTR szHost
				, LPCTSTR szProgID
				, double dConnTime
//				, bool& bResult
					 )
{

/*
	int curIndex=0;			// Индекс элемента выбранного в box
	sCurrStep = _T("Инициализация дескриптора окружения");
	m_ListConnect.InsertItem(m_ListConnect.GetItemCount(), sSuccses+sCurrStep);
//	m_ListConnect.RedrawItems(0,nColumnCount);
//	m_ListConnect.DeleteAllItems();

//	m_ListConnect.InsertColumn(0, _T("Статус подключения"), LVCFMT_LEFT, 1000);
	//m_ListConnect.InsertItem(0, _T("1"));
	//m_ListResult.SetItemText(i, j, v_SQLFullTable[i][j]);
*/
	bNextStep=false;

	// have to be done before using microsoft COM library:
	CurMethod = _T("CoInitialize");
	hr = CoInitialize(NULL);
	if(hr==S_OK)
	{AddItemToListBox(_T("ОК - Загрузилась библиотека СОМ"), 1);}
	else
	{
		AddItemToListBox(_T("ОШИБКА - Не загрузилась библиотека СОМ"), 1);
		return false;
	}

	pIOPCServer = NULL;		//pointer to IOPServer interface
	pIOPCItemMgt = NULL;	//pointer to IOPCItemMgt interface

	CurMethod = _T("InstantiateServer");
	AddItemToListBox(CurMethod, 1);
	pIOPCServer = InstantiateServer(szProgID, szHost, bNextStep);
	//Sleep(10000);

	// Add the OPC group the OPC server and get an handle to the IOPCItemMgt interface:
	if(bNextStep)
	{
		//return _T("");
		AddItemToListBox(_T("OK - ") + CurMethod, 1);
		return true;
	}
	else
	{
		//return FullErrMess;
		AddItemToListBox(_T("ОШИБКА - ") + CurMethod, 1);
		return false;
	}
}


bool CMFC_OPC_ExplorerDlg::DisConnect()
{
	CurMethod = _T("DisConnect");
	bool bIOPCServer(false), bIOPCItemMgt(false);

	if((pIOPCItemMgt!=NULL) && pIOPCItemMgt)
	{
		AddItemToListBox(_T("OK - Release( pIOPCItemMgt )"), 1);

		pIOPCItemMgt->Release();
		pIOPCItemMgt=NULL;
		bIOPCItemMgt=true;
	}

	if((pIOPCServer!=NULL) && pIOPCServer)
	{
		AddItemToListBox(_T("OK - Release ( pIOPCServer )"), 1);

		pIOPCServer->Release();
		pIOPCServer=NULL;
		bIOPCServer=true;
	}
	

	//err_wf.Close();
	//res_wf.Close();

	CString NoParams("");
	LogSessionStatus(_T("Сессия закрыта"), NoParams, NoParams, false);

	CoUninitialize();

	if(bIOPCServer)
	{
		AddItemToListBox(_T("OK - ") + CurMethod, 1);
		return true;
	}
	else
	{
		AddItemToListBox(_T("ОШИБКА - ") + CurMethod, 1);
		return false;
	}
}





///////////////////////////////////////////////////////////////////////////////////////
// МЕТОДЫ РЕАЛИЗАЦИЯ И Т.Д. ///////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////
// Instantiate the IOPCServer interface of the OPCServer
// having the name ServerName. Return a pointer to this interface
//
//IOPCServer* InstantiateServer(wchar_t OPCServerName[], wchar_t RemoteServerName[])
IOPCServer* CMFC_OPC_ExplorerDlg::InstantiateServer(CString OPCServProgID_CStr, CString RmtServNameOrIPAdr_CStr, bool& Result)
{
	//bool Result=true;
	CLSID CLSID_OPCServer;
	HRESULT hr(NULL);

	wchar_t *OPCServProgID_Wcht;
	OPCServProgID_Wcht = (wchar_t *)(LPCTSTR)(LPCTSTR)OPCServProgID_CStr;
	wchar_t *RmtServNameOrIPAdr_Wcht;
	RmtServNameOrIPAdr_Wcht = (wchar_t *)(LPCTSTR)(LPCTSTR)RmtServNameOrIPAdr_CStr;


	CurMethod = _T("CLSIDFromProgID");
	CurParams.Format(_T("ProgID=%s"), OPCServProgID_CStr);


	// get the CLSID from the OPC Server Name:
	hr = CLSIDFromString(OPCServProgID_Wcht, &CLSID_OPCServer);
	if(hr != S_OK) //if(FAILED(hr))
	{
		CurErr = GetErrorCSTRString(hr);
		DeleteLastSymbols(CurErr, 2);
		long lhr = HRESULT_CODE(hr);
		CurErrCode.Format(_T("%d"),lhr);
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
		//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
		//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);
		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);

		AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 1);

		Result = false;

		/*
		ErrorString = GetErrorString(hr);
		wcout <<ErrorString<<endl;

		ErrMess = ErrMess + _T("Failed  - CLSIDFromProgID()") + _T("\n")
						  + _T("with Parametrs:") + _T("\n")
						  + _T("OPCServProgID = ") + OPCServProgID_CStr + _T("\n");

		CString sErrorMsg = _com_error(hr).ErrorMessage();
		ErrMess = ErrMess + _T("Error:") + _T("\n")
						  //+ sErrorMsg + _T("\n");
						  + ErrorString.c_str() + _T("\n");

		ErrorString.clear();
		//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);	//_ASSERT(!FAILED(hr));
		Result = false;
		//*/
	}
	else
	{
		AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 1);
//		cout <<"Success - Method CLSIDFromProgID()"<<endl;
		Result = true;
	}

////	const IID IID_IOPCServer;

	//queue of the class instances to create
	LONG cmq = 1; // number of class instance to create.
	MULTI_QI queue[1] =
		{{&IID_IOPCServer
		, NULL
		, 0
		}};

	if(RmtServNameOrIPAdr_CStr.GetLength() > 0)
	{
////////////////////////////////////////////////////////////////////
// 2 // Подключение к OPC-серверу в одной из удалённых сетей
/*
		COSERVERINFO csi = {0};	// null out all fields.

		// Set the name of the remote server.
		csi.pwszName = (L"BIGMANU"); // ADD YOUR SERVER NAME HERE!!!!!!!! 

		// Here are the interfaces I want.
		MULTI_QI qi[3] = {0};
		qi[0].pIID = &IID_IEngine;
		qi[1].pIID = &IID_IStats;
		qi[2].pIID = &IID_ICreateCar;

		// Create the remote CoCar.
		hr = CoCreateInstanceEx(CLSID_CoCar, NULL, CLSCTX_REMOTE_SERVER, &csi, 3, qi);

		// Assign interface pointers to fetched results.
		IEngine *    pEngine = (IEngine *   ) qi[0].pItf;
		IStats *     pStats  = (IStats *    ) qi[1].pItf;
		ICreateCar * pCCar   = (ICreateCar *) qi[2].pItf;
//*/
		if(Result)
		{
			//Server info:
			COSERVERINFO CoServerInfo =
			{
				0 // dwReserved1
		//		,REMOTE_SERVER_NAME //pwszName
				,RmtServNameOrIPAdr_Wcht //pwszName
				,NULL //COAUTHINFO
				,0 //dwReserved2
			}; 

			CurMethod = _T("CoCreateInstanceEx");
			CurParams.Format(_T("ProgID=%s, Host=%s"), OPCServProgID_CStr, RmtServNameOrIPAdr_CStr);

			hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER,
				&CoServerInfo,
				cmq, queue);
			
			if(hr != S_OK) //if(FAILED(hr))
			{
				CurErr = GetErrorCSTRString(hr);
				DeleteLastSymbols(CurErr, 2);
				long lhr = HRESULT_CODE(hr);
				CurErrCode.Format(_T("%d"),lhr);
				CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

				AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 1);

				LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
				//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
				//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
				//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);

				Result = false;

				/*
				ErrorString = GetErrorCSTRString(hr);
				wcout <<ErrorString<<endl;


				ErrMess = ErrMess + _T("Failed  - CoCreateInstanceEx()") + _T("\n")
								  + _T("with Parametrs:") + _T("\n")
								  + _T("OPCServProgID = ") + OPCServProgID_CStr + _T("\n");
								  + _T("HostName = ") + RemServNameOrIPAdr_CStr + _T("\n");

				CString sErrorMsg = _com_error(hr).ErrorMessage();
				ErrMess = ErrMess + _T("Error:") + _T("\n")
								  //+ sErrorMsg + _T("\n");
								  + ErrorString.c_str() + _T("\n");

				ErrorString.clear();
				//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);	// _ASSERT(!hr);
				Result = false;
				*/
			}
			else
			{
				AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 1);
//				cout <<"Success - Method CoCreateInstanceEx()"<<endl;
			}
		}

// 2 // Подключение Удалённое / к OPC-серверу в одной из удалённых сетей
////////////////////////////////////////////////////////////////////
	}
	else
	{
////////////////////////////////////////////////////////////////////
// 1 // Подключение напрямую / Локальное
		if(Result)
		{
			CurMethod = _T("CoCreateInstanceEx");
			CurParams.Format(_T("ProgID=%s"), OPCServProgID_CStr);

			// create an instance of the IOPCServer
			hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER,
				NULL, // &CoServerInfo
				cmq, queue);
			if(hr != S_OK) //if(FAILED(hr))
			{
				CurErr = GetErrorCSTRString(hr);
				DeleteLastSymbols(CurErr, 2);
				long lhr = HRESULT_CODE(hr);
				CurErrCode.Format(_T("%d"),lhr);
				CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

				AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 1);

				LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);

				//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
				//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
				//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);

				Result = false;

				/*
				ErrorString = GetErrorCSTRString(hr);
				wcout <<ErrorString<<endl;

				ErrMess = ErrMess + _T("Failed  - CoCreateInstanceEx()") + _T("\n")
								  + _T("with Parametrs:") + _T("\n")
								  + _T("OPCServProgID = ") + OPCServProgID_CStr + _T("\n");
								  + _T("HostName = NULL (Local)\n");

				CString sErrorMsg = _com_error(hr).ErrorMessage();
				ErrMess = ErrMess + _T("Error:") + _T("\n")
								  //+ sErrorMsg + _T("\n");
								  + ErrorString.c_str() + _T("\n");

				ErrorString.clear();
				//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);	// _ASSERT(!hr);
				Result = false;
				*/
			}
			else
			{
				AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 1);
//				cout <<"Success - Method CoCreateInstanceEx()"<<endl;
			}
		}
// 1 // Подключение напрямую / Локальное
////////////////////////////////////////////////////////////////////
	}

	
	if(Result)
	{
//		cout <<"Success - Method InstantiateServer()"<<endl;
		// return a pointer to the IOPCServer interface:
		return(IOPCServer*) queue[0].pItf;
	}
	else
	{
		return(IOPCServer*) NULL;
		//return NULL;
	}
}


/////////////////////////////////////////////////////////////////////
// Add group "Group1" to the Server whose IOPCServer interface
// is pointed by pIOPCServer. 
// Returns a pointer to the IOPCItemMgt interface of the added group
// and a server opc handle to the added group.
//
//void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, OPCHANDLE& hServerGroup)
bool CMFC_OPC_ExplorerDlg::AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, OPCHANDLE& hServerGroup)
{
	bool Result=true;

	DWORD dwUpdateRate = 0;
	OPCHANDLE hClientGroup = 0;

////	const IID IID_IOPCItemMgt;

	//CurMethod = _T("AddGroup");
	//CurParams = _T("GrName=Grp1,...");
	
	CurMethod = _T("AddGroup");
	CurParams = _T("GrName = ") + sGroupTagName;

	// Add an OPC group and get a pointer to the IUnknown I/F:
	HRESULT hr = pIOPCServer->AddGroup(
		sGroupTagName
		//L"Grp1"		// szName
		//,FALSE		// bActive
		,TRUE			// bActive
		,dwUpdateRate	// dwRequestedUpdateRate
		,hClientGroup	// hClientGroup
		,0				// pTimeBias
		,0				// pPercentDeadband
		,0				// dwLCID
		,&hServerGroup	// phServerGroup
		,&dwUpdateRate
		,IID_IOPCItemMgt //riid
		,(IUnknown**) &pIOPCItemMgt //ppUnk
		);
		
	if(FAILED(hr))
	{
		CurErr = GetErrorCSTRString(hr);
		DeleteLastSymbols(CurErr, 2);
		long lhr = HRESULT_CODE(hr);
		CurErrCode.Format(_T("%d"),lhr);
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
		//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
		//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);

		Result = false;

		/*
		ErrorString = GetErrorCSTRString(hr);
		wcout <<ErrorString<<endl;

		ErrMess = ErrMess + _T("Failed  - AddGroup()") + _T("\n")
						  + _T("with Parametrs:") + _T("\n")
						  + _T("szName = Group1") + _T("\n");

		CString sErrorMsg = _com_error(hr).ErrorMessage();
		ErrMess = ErrMess + _T("Error:") + _T("\n")
						  //+ sErrorMsg + _T("\n");
						  + ErrorString.c_str() + _T("\n");

		ErrorString.clear();
		//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);	// 	_ASSERT(!FAILED(hr));
		return false;
		*/
	}
	else
	{
		AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);
//		cout <<"Success - Method AddGroup()"<<endl;
		Result = true;
	}

	return Result;
}




//////////////////////////////////////////////////////////////////
// Add the Item ITEM_ID to the group whose IOPCItemMgt interface
// is pointed by pIOPCItemMgt pointer. Return a server opc handle
// to the item.

//void AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, int iItemsNum, std::vector<CString> v_AllTags)
int CMFC_OPC_ExplorerDlg::AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, int iItemsNum, std::vector<std::vector<CString>> v_AllTags)
{
	int Result(0);

	HRESULT hr(NULL);

/*	// Array of items to add:
	OPCITEMDEF ItemArray[1] =
	{{
	L"",		// szAccessPath
	ITEM_ID,	// szItemID
	FALSE,		// bActive
	1,			// hClient
	0,			// dwBlobSize
	NULL,		// pBlob
	VT,			// vtRequestedDataType
	0			//wReserved
	}};
//*/

	// Array of items to add:
//	OPCITEMDEF ItemArray[TAGS_NUM]; //#define TAGS_NUM 5

/*
//#define VCTR_INDX_ACTVSOU 0	// Активен/нет/СОУ
#define VCTR_INDX_TAGNAME 0	// Имя тега
#define VCTR_INDX_TAGTYPE 1	// Тип тега
#define VCTR_INDX_INVRS	  2	// Тег инвертирован
*/

	LPWSTR ItemID;
	ItemID = (LPWSTR)(LPCTSTR)v_AllTags[iItemsNum][VCTR_INDX_TAGNAME]; // (LPSTR)(LPCSTR)

	CString VariantTYPE_CStr("");
	VariantTYPE_CStr = (LPWSTR)(LPCTSTR)v_AllTags[iItemsNum][VCTR_INDX_TAGTYPE];

	VARTYPE VariantTYPE;
	VariantTYPE = CStrVarTypeToVariantType(VariantTYPE_CStr);


	OPCITEMDEF ItemArray[1]=
	{{
	L"",		 // szAccessPath
	//ITEM_ID,	 // szItemID
	ItemID,
	//FALSE,	 // bActive // В случае false потребуется вызвать после метод SetState
	TRUE,		 // bActive
	1,			 // hClient
	0,			 // dwBlobSize
	NULL,		 // pBlob
	//DEFINED_VT,// vtRequestedDataType
	VariantTYPE,
	0			 //wReserved
	}};

	//Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;


	CurMethod = _T("AddItems");
	CurParams.Format(_T("%s(%s)"), ItemID, VariantTYPE_CStr); //iItemsNum

	// Add an Item to the previous Group:
//	hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);

//*
//////////////////////////////////////////////////////////////////
	time(&currentTime);	startTime = currentTime;
			
	// read the item value from the device:
	hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);

	time(&currentTime);	timeElapsed = currentTime - startTime;
//////////////////////////////////////////////////////////////////

	const int iTime(19);
	if(timeElapsed <= iTime)//theApp.Val_TagsTimeout)
	{Result=0;}
	else
	{
		CurErr.Format(_T("Истекло время ожидания запроса Тега в %d секунд. Связь потеряна!"), iTime); //theApp.Val_TagsTimeout);
		CurErrCode.Format(_T("0"));
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		Result=1;
	}
//*/


	if(!Result)
	{
		if(hr != S_OK) //if(FAILED(hr))
		{
			CurErr = GetErrorCSTRString(hr);
			DeleteLastSymbols(CurErr, 2);
			long lhr = HRESULT_CODE(hr);
			CurErrCode.Format(_T("%d"),lhr);
			CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

			AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

			LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
			//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
			//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
			//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);

			Result = 2;

			/*
			ErrorString = GetErrorCSTRString(hr);
			wcout <<ErrorString<<endl;

			ErrMess = ErrMess + _T("Failed  - AddItems()") + _T("\n")
							  + _T("with Parametrs:") + _T("\n")
							  + _T("ItemArray[1] -> ItemID = ") + v_AllTags[iItemsNum] + _T("\n")
							  + _T("ItemArray[1] -> DataType = VT_BOOL") + _T("\n");

			//CString sErrorMsg = _com_error(hr).ErrorMessage();
			CString sErrorMsg = _com_error(*pErrors).ErrorMessage();
			ErrMess = ErrMess + _T("Error:") + _T("\n")
							  //+ sErrorMsg + _T("\n");
							  + ErrorString.c_str() + _T("\n");

			ErrorString.clear();
			//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);
			return false;
			*/
		}
		else
		{
			AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);

	//		cout <<"Success - Method AddItems()"<<endl;
			//_ASSERT(!FAILED(hr));	//_ASSERT(!hr);

			// Server handle for the added item:
			hServerItem = pAddResult[0].hServer;

			// release memory allocated by the server:
			CoTaskMemFree(pAddResult->pBlob);
			CoTaskMemFree(pAddResult);
			pAddResult = NULL;

			CoTaskMemFree(pErrors);
			pErrors = NULL;

			Result = 0;
		}
	}


	return Result;
}


//////////////////////////////////////////////////////////////////
// Add the Item ITEM_ID to the group whose IOPCItemMgt interface
// is pointed by pIOPCItemMgt pointer. Return a server opc handle
// to the item.

bool CMFC_OPC_ExplorerDlg::AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, CString TagName, CString TagType)
{
	bool Result(true);

	HRESULT hr(NULL);

	LPWSTR ItemID;
	ItemID = (LPWSTR)(LPCTSTR)TagName; // (LPSTR)(LPCSTR)

	VARTYPE VAR_TYPE = CStrVarTypeToVariantType(TagType);

	OPCITEMDEF ItemArray[1]=
	{{
	L"",		// szAccessPath
//	ITEM_ID,	// szItemID
	ItemID,
//	FALSE,		// bActive // В случае false потребуется вызвать после метод SetState
	TRUE,		// bActive
	1,			// hClient
	0,			// dwBlobSize
	NULL,		// pBlob
	VAR_TYPE,	// vtRequestedDataType
	0			//wReserved
	}};

	//Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;

	CurMethod = _T("AddItems");
	CurParams.Format(_T("s%(VT_BOOL)"), ItemID); //iItemsNum

	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);
	if(hr != S_OK) //if(FAILED(hr))
	{
		CurErr = GetErrorCSTRString(hr);
		DeleteLastSymbols(CurErr, 2);
		long lhr = HRESULT_CODE(hr);
		CurErrCode.Format(_T("%d"),lhr);
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
		//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
		//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);
	
		Result = false;

		/*
		ErrorString = GetErrorCSTRString(hr);
		wcout <<ErrorString<<endl;

		ErrMess = ErrMess + _T("Failed  - AddItems()") + _T("\n")
						  + _T("with Parametrs:") + _T("\n")
						  + _T("ItemArray[1] -> ItemID = ") + v_AllTags[iItemsNum] + _T("\n")
						  + _T("ItemArray[1] -> DataType = VT_BOOL") + _T("\n");

		//CString sErrorMsg = _com_error(hr).ErrorMessage();
		CString sErrorMsg = _com_error(*pErrors).ErrorMessage();
		ErrMess = ErrMess + _T("Error:") + _T("\n")
						  //+ sErrorMsg + _T("\n");
						  + ErrorString.c_str() + _T("\n");

		ErrorString.clear();
		//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);
		return false;
		*/
	}
	else
	{
//		cout <<"Success - Method AddItems()"<<endl;
		//_ASSERT(!FAILED(hr));	//_ASSERT(!hr);

		// Server handle for the added item:
		hServerItem = pAddResult[0].hServer;

		// release memory allocated by the server:
		CoTaskMemFree(pAddResult->pBlob);
		CoTaskMemFree(pAddResult);
		pAddResult = NULL;

		CoTaskMemFree(pErrors);
		pErrors = NULL;

		Result = true;
	}

	return Result;
}




int CMFC_OPC_ExplorerDlg::AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, TagStruct OneTag)
{
	int Result(0);

	HRESULT hr(NULL);

/*	// Array of items to add:
	OPCITEMDEF ItemArray[1] =
	{{
	  L""		// szAccessPath
	, ITEM_ID	// szItemID
	, FALSE		// bActive
	, 1			// hClient
	, 0			// dwBlobSize
	, NULL		// pBlob
	, VT		// vtRequestedDataType
	, 0			//wReserved
	}};
//*/

	LPWSTR ItemID;
	ItemID = (LPWSTR)(LPCTSTR)OneTag.TagFullName; // (LPSTR)(LPCSTR)

	CString VariantTYPE_CStr("");
	VariantTYPE_CStr = VariantTypeToCStrVarType(OneTag.TagVrntType);

	VARTYPE VariantTYPE;
	VariantTYPE = OneTag.TagVrntType;

	OPCITEMDEF ItemArray[1]=
	{{
	  L""			// szAccessPath
//	, ITEM_ID		// szItemID
	, ItemID
//	, FALSE			// bActive // В случае false потребуется вызвать после метод SetState
	, TRUE			// bActive
	, 1				// hClient
	, 0				// dwBlobSize
	, NULL			// pBlob
	, VariantTYPE	
	, 0				//wReserved
	}};

	//Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;


	CurMethod = _T("AddItems");
	CurParams.Format(_T("%s(%s)"), ItemID, VariantTYPE_CStr); //iItemsNum

	// Add an Item to the previous Group:
	//hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);

//*
//////////////////////////////////////////////////////////////////
	time(&currentTime);	startTime = currentTime;
			
	// read the item value from the device:
	hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);

	time(&currentTime);	timeElapsed = currentTime - startTime;
//////////////////////////////////////////////////////////////////

	const int iTime(19);
	if(timeElapsed <= iTime)//theApp.Val_TagsTimeout)
	{Result=0;}
	else
	{
		CurErr.Format(_T("Истекло время ожидания запроса Тега в %d секунд. Связь потеряна!"), iTime); //theApp.Val_TagsTimeout);
		CurErrCode.Format(_T("0"));
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		Result=1;
	}
//*/


	if(!Result)
	{
		if(hr != S_OK) //if(FAILED(hr))
		{
			CurErr = GetErrorCSTRString(hr);
			DeleteLastSymbols(CurErr, 2);
			long lhr = HRESULT_CODE(hr);
			CurErrCode.Format(_T("%d"),lhr);
			CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

			AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

			LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
			//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
			//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
			//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);

			Result = 2;

			/*
			ErrorString = GetErrorCSTRString(hr);
			wcout <<ErrorString<<endl;

			ErrMess = ErrMess + _T("Failed  - AddItems()") + _T("\n")
							  + _T("with Parametrs:") + _T("\n")
							  + _T("ItemArray[1] -> ItemID = ") + v_AllTags[iItemsNum] + _T("\n")
							  + _T("ItemArray[1] -> DataType = VT_BOOL") + _T("\n");

			//CString sErrorMsg = _com_error(hr).ErrorMessage();
			CString sErrorMsg = _com_error(*pErrors).ErrorMessage();
			ErrMess = ErrMess + _T("Error:") + _T("\n")
							  //+ sErrorMsg + _T("\n");
							  + ErrorString.c_str() + _T("\n");

			ErrorString.clear();
			//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);
			return false;
			*/
		}
		else
		{
			AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);

	//		cout <<"Success - Method AddItems()"<<endl;
			//_ASSERT(!FAILED(hr));	//_ASSERT(!hr);

			// Server handle for d added item:
			hServerItem = pAddResult[0].hServer;

			// release memory allocated by the server:
			CoTaskMemFree(pAddResult->pBlob);
			CoTaskMemFree(pAddResult);
			pAddResult = NULL;

			CoTaskMemFree(pErrors);
			pErrors = NULL;

			Result = 0;
		}
	}


	return Result;
}




///////////////////////////////////////////////////////////////////////////////
// Read from device the value of the item having the "hServerItem" server 
// handle and belonging to the group whose one interface is pointed by
// pGroupIUnknown. The value is put in varValue. 
//
//void ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
//bool CMFC_OPC_ExplorerDlg::ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
bool CMFC_OPC_ExplorerDlg::ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue, WORD& wrdValue)
{
	bool Result=true;

	// value of the item:
	OPCITEMSTATE* pValue = NULL;

	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);


	CurMethod = _T("ReadItem");
	CurParams.Format(_T("%s(%s)"), sTagName, sTagType);
	//CurParams = VariantToCStr(varValue);
	HRESULT* pErrors = NULL; //to store error code(s)

	HRESULT hr = NULL;
	
	int CheckRadio = GetCheckedRadioButton(IDC_BTN_TAG_READ_CACHE, IDC_BTN_TAG_READ_DEVICE);
	
	if(CheckRadio == IDC_BTN_TAG_READ_CACHE)
	{
	 hr = pIOPCSyncIO->Read(
							//  OPC_DS_DEVICE	// ЧТЕНИЕ из УСТРОЙСТВА на прямую
							  OPC_DS_CACHE		// ЧТЕНИЕ из КЭША
							, 1
							, &hServerItem
							, &pValue
							, &pErrors
							);
	}

	if(CheckRadio == IDC_BTN_TAG_READ_DEVICE)
	{
	 hr = pIOPCSyncIO->Read(
							  OPC_DS_DEVICE	// ЧТЕНИЕ из УСТРОЙСТВА на прямую
							//  OPC_DS_CACHE		// ЧТЕНИЕ из КЭША
							, 1
							, &hServerItem
							, &pValue
							, &pErrors
							);
	}


/*
//////////////////////////////////////////////////////////////////
	time(&currentTime);	startTime = currentTime;
			
	// read the item value from the device:
	HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);

	time(&currentTime);	timeElapsed = currentTime - startTime;
//////////////////////////////////////////////////////////////////


	if(timeElapsed <= theApp.Val_TagsTimeout)
	{Result=true;}
	else
	{
		CurErr.Format(_T("Истекло время ожидания запроса Тега в %d секунд. Связь потеряна!"), theApp.Val_TagsTimeout);
		CurErrCode.Format(_T("0"));
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		Result=false;
	}
*/
	if(Result)
	{
		if( (hr != S_OK) || (pValue==NULL) )
		{
			CurErr = GetErrorCSTRString(hr);
			DeleteLastSymbols(CurErr, 2);
			long lhr = HRESULT_CODE(hr);
			CurErrCode.Format(_T("%d"),lhr);
			CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

			AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

			LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
			//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
			//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
			//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);

			Result=false;

			/*
			ErrorString = GetErrorCSTRString(hr);
			wcout <<ErrorString<<endl;

			ErrMess = ErrMess + _T("Failed  - Read()") + _T("\n");

			//CString sErrorMsg = _com_error(hr).ErrorMessage();
			CString sErrorMsg = _com_error(*pErrors).ErrorMessage();
			ErrMess = ErrMess + _T("Error:") + _T("\n")
							//  + sErrorMsg + _T("\n");
								+ ErrorString.c_str() + _T("\n");

			ErrorString.clear();
			//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);
			return false;
			*/
		}
		else
		{
			AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);

	//		cout <<"Success - Method ReadItem()"<<endl;
			//_ASSERT(!hr);
			//_ASSERT(pValue!=NULL);

			varValue = pValue[0].vDataValue;
			wrdValue = pValue[0].wQuality;

			//Release memeory allocated by the OPC server:
			CoTaskMemFree(pErrors);
			pErrors = NULL;

			CoTaskMemFree(pValue);
			pValue = NULL;

			// release the reference to the IOPCSyncIO interface:
			pIOPCSyncIO->Release();
		}
	}

	return Result;
}


bool CMFC_OPC_ExplorerDlg::ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, TagStruct& OneTag)
{
	bool Result=true;

	// value of the item:
	OPCITEMSTATE* pValue = NULL;

	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);


	CurMethod = _T("ReadItem");
	CurParams.Format(_T("%s(%s)"), OneTag.TagFullName, VariantTypeToCStrVarType(OneTag.TagVrntType));
	//CurParams = VariantToCStr(varValue);
	HRESULT* pErrors = NULL; //to store error code(s)
	

	///////////////////////////////////////////////
	// ЧТЕНИЕ из УСТРОЙСТВА на прямую OPC_DS_DEVICE
	// HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);


	///////////////////////////////////////////////
	// ЧТЕНИЕ из КЭША на прямую OPC_DS_CACHE
	 HRESULT hr = pIOPCSyncIO->Read(OPC_DS_CACHE, 1, &hServerItem, &pValue, &pErrors);


/*
//////////////////////////////////////////////////////////////////
	time(&currentTime);	startTime = currentTime;
			
	// read the item value from the device:
	HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);

	time(&currentTime);	timeElapsed = currentTime - startTime;
//////////////////////////////////////////////////////////////////


	if(timeElapsed <= theApp.Val_TagsTimeout)
	{Result=true;}
	else
	{
		CurErr.Format(_T("Истекло время ожидания запроса Тега в %d секунд. Связь потеряна!"), theApp.Val_TagsTimeout);
		CurErrCode.Format(_T("0"));
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		Result=false;
	}
*/
	if(Result)
	{
		if( (hr != S_OK) || (pValue==NULL) )
		{
			CurErr = GetErrorCSTRString(hr);
			DeleteLastSymbols(CurErr, 2);
			long lhr = HRESULT_CODE(hr);
			CurErrCode.Format(_T("%d"),lhr);
			CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

			AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

			LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
			//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
			//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
			//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);

			Result=false;

			/*
			ErrorString = GetErrorCSTRString(hr);
			wcout <<ErrorString<<endl;

			ErrMess = ErrMess + _T("Failed  - Read()") + _T("\n");

			//CString sErrorMsg = _com_error(hr).ErrorMessage();
			CString sErrorMsg = _com_error(*pErrors).ErrorMessage();
			ErrMess = ErrMess + _T("Error:") + _T("\n")
							//  + sErrorMsg + _T("\n");
								+ ErrorString.c_str() + _T("\n");

			ErrorString.clear();
			//::MessageBox(NULL, ErrMess, _T("Error"), 0x10010);
			return false;
			*/
		}
		else
		{
			AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);

	//		cout <<"Success - Method ReadItem()"<<endl;
			//_ASSERT(!hr);
			//_ASSERT(pValue!=NULL);

			//varValue = pValue[0].vDataValue;
			//wrdValue = pValue[0].wQuality;
			OneTag.TagVrntValue	= pValue[0].vDataValue;
			OneTag.TagVrntType	= pValue[0].vDataValue.vt;
			OneTag.TagQuality	= pValue[0].wQuality;
			OneTag.TagTimeStamp	= pValue[0].ftTimeStamp;
//			OneTag.TagHandle	= pValue[0].hClient;

			OneTag.TagCStrValue = VariantValueToCStr(OneTag.TagVrntValue);

			//Release memeory allocated by the OPC server:
			CoTaskMemFree(pErrors);
			pErrors = NULL;

			CoTaskMemFree(pValue);
			pValue = NULL;

			// release the reference to the IOPCSyncIO interface:
			pIOPCSyncIO->Release();
		}
	}

	return Result;
}



bool CMFC_OPC_ExplorerDlg::WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT* varValue)
{
	bool Result=true;
	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	IOPCAsyncIO* pIOPCAsyncIO;

	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCAsyncIO), (void**) &pIOPCAsyncIO);

	// read the item value from the device:
	HRESULT hr(NULL);
	HRESULT hRes_Mess(NULL);
	HRESULT* pErrors(NULL); //to store error code(s)

//	DWORD   dwConnection;
	OPCHANDLE  phServer[1]; // server handle of our item
	VARIANT   ItemValues[1]; // item value
//	DWORD   TransactionID[1];

	//HRESULT*  pErrors;  // read errors
	//IDataObject     pID
	phServer[0] = hServerItem;
	ItemValues[0].bVal = 1;
  

	CurMethod = _T("WriteItem");
	CurParams.Format(_T("%s(%s)"), sTagName, sTagType);
	//CurParams = VariantToCStr(*varValue);

	//HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);
	hr = pIOPCSyncIO->Write(1, &hServerItem, varValue, &pErrors);

	//wcout <<"Method WriteItem():"<<endl;
	//ErrorString = GetErrorCSTRString(hr);
	//wcout <<ErrorString<<endl;
	//ErrorString.clear();

	hRes_Mess = *pErrors;
	//hr = pIOPCAsyncIO->Write(dwConnection,1,phServer,ItemValues,TransactionID,&pErrors);

	//_com_error err(hr);
	//LPCTSTR errMsg = err.ErrorMessage();

	LPTSTR errorText = NULL;

	FormatMessage(
	   // use system message tables to retrieve error text
	   FORMAT_MESSAGE_FROM_SYSTEM
	   // allocate buffer on local heap for error text
	   |FORMAT_MESSAGE_ALLOCATE_BUFFER
	   // Important! will fail otherwise, since we're not (and CANNOT) pass insertion parameters
	   |FORMAT_MESSAGE_IGNORE_INSERTS,  
	   NULL,    // unused with FORMAT_MESSAGE_FROM_SYSTEM
	   hRes_Mess,
	   MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
	   (LPTSTR)&errorText,  // output 
	   0, // minimum size for output buffer
	   NULL);   // arguments - see note 

	if( (hr!=S_OK)&&( NULL != errorText ) )
	{
		//AfxMessageBox(errorText);
	   // ... do something with the string - log it, display it to the user, etc.
		//CurErr = GetErrorCSTRString(hr);
		//DeleteLastSymbols(CurErr, 2);
		CurErr = (LPCTSTR)errorText;
		long lhr = HRESULT_CODE(hr);
		CurErrCode.Format(_T("%d"),lhr);
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
		//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
		//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);
		Result=false;

	   // release memory allocated by FormatMessage()
	   LocalFree(errorText);
	   errorText = NULL;
	}	
	else
	{
		AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);
		//::MessageBox(NULL, CurMethod, _T("Success"), MB_OK);	// _ASSERT(!hr);

		//Release memeory allocated by the OPC server:
		CoTaskMemFree(pErrors);
		pErrors = NULL;

		// release the reference to the IOPCSyncIO interface:
		pIOPCSyncIO->Release();
	}

	return Result;
}




bool CMFC_OPC_ExplorerDlg::WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, TagStruct OneTag)
{
	bool Result=true;
	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	IOPCAsyncIO* pIOPCAsyncIO;

	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCAsyncIO), (void**) &pIOPCAsyncIO);

	// read the item value from the device:
	HRESULT hr(NULL);
	HRESULT hRes_Mess(NULL);
	HRESULT* pErrors(NULL); //to store error code(s)

//	DWORD   dwConnection;
	OPCHANDLE  phServer[1]; // server handle of our item
	VARIANT   ItemValues[1]; // item value
//	DWORD   TransactionID[1];

	//HRESULT*  pErrors;  // read errors
	//IDataObject     pID
	phServer[0] = hServerItem;
	ItemValues[0].bVal = 1;
  

	CurMethod = _T("WriteItem");
	CurParams.Format(_T("%s(%s)"), sTagName, sTagType);
	//CurParams = VariantToCStr(*varValue);

	//HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);
	hr = pIOPCSyncIO->Write(1, &hServerItem, &OneTag.TagVrntValue, &pErrors);

	//wcout <<"Method WriteItem():"<<endl;
	//ErrorString = GetErrorCSTRString(hr);
	//wcout <<ErrorString<<endl;
	//ErrorString.clear();

	hRes_Mess = *pErrors;
	//hr = pIOPCAsyncIO->Write(dwConnection,1,phServer,ItemValues,TransactionID,&pErrors);

	//_com_error err(hr);
	//LPCTSTR errMsg = err.ErrorMessage();

	LPTSTR errorText = NULL;

	FormatMessage(
	   // use system message tables to retrieve error text
	   FORMAT_MESSAGE_FROM_SYSTEM
	   // allocate buffer on local heap for error text
	   |FORMAT_MESSAGE_ALLOCATE_BUFFER
	   // Important! will fail otherwise, since we're not (and CANNOT) pass insertion parameters
	   |FORMAT_MESSAGE_IGNORE_INSERTS,  
	   NULL,    // unused with FORMAT_MESSAGE_FROM_SYSTEM
	   hRes_Mess,
	   MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
	   (LPTSTR)&errorText,  // output 
	   0, // minimum size for output buffer
	   NULL);   // arguments - see note 

	if( (hr!=S_OK)&&( NULL != errorText ) )
	{
		//AfxMessageBox(errorText);
	   // ... do something with the string - log it, display it to the user, etc.
		//CurErr = GetErrorCSTRString(hr);
		//DeleteLastSymbols(CurErr, 2);
		CurErr = (LPCTSTR)errorText;
		long lhr = HRESULT_CODE(hr);
		CurErrCode.Format(_T("%d"),lhr);
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		//FullErrMess = GetFullErrMess(CurMethod, CurParams, CurErr, CurErrCode);
		//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
		//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);
		Result=false;

	   // release memory allocated by FormatMessage()
	   LocalFree(errorText);
	   errorText = NULL;
	}	
	else
	{
		AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);
		//::MessageBox(NULL, CurMethod, _T("Success"), MB_OK);	// _ASSERT(!hr);

		//Release memeory allocated by the OPC server:
		CoTaskMemFree(pErrors);
		pErrors = NULL;

		// release the reference to the IOPCSyncIO interface:
		pIOPCSyncIO->Release();
	}

	return Result;
}


/*
CString FormatMessage(HRESULT result)
{
    CString strMessage;
    WORD facility = HRESULT_FACILITY(result);
    CComPtr<IErrorInfo> iei;
    if (S_OK == GetErrorInfo(0, &iei) && iei)
    {
        // get the error description from the IErrorInfo 
        BSTR bstr = NULL;
        if (SUCCEEDED(iei->GetDescription(&bstr)))
        {
            // append the description to our label
            strMessage.Append(bstr);

            // done with BSTR, do manual cleanup
            SysFreeString(bstr);
        }
    }
    else if (facility == FACILITY_ITF)
    {
        // interface specific - no standard mapping available
        strMessage.Append(_T("FACILITY_ITF - This error is interface specific.  No further information is available."));
    }
    else
    {
        // attempt to treat as a standard, system error, and ask FormatMessage to explain it
        CString error;
        //CErrorMessage::FormatMessage(error, result); // <- This is just a wrapper for ::FormatMessage, left to reader as an exercise :)
        FormatMessage(error, result); // <- This is just a wrapper for ::FormatMessage, left to reader as an exercise :)
        if (!error.IsEmpty())
            strMessage.Append(error);
    }
    return strMessage;
}
*/

///////////////////////////////////////////////////////////////////////////
// Remove the item whose server handle is hServerItem from the group
// whose IOPCItemMgt interface is pointed by pIOPCItemMgt
//
void CMFC_OPC_ExplorerDlg::RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE hServerItem)
{
	// server handle of items to remove:
	OPCHANDLE hServerArray[1];
	hServerArray[0] = hServerItem;
	
	//Remove the item:
	CurMethod = _T("RemoveItem");
	CurParams = _T("hServerArray");

	HRESULT* pErrors=NULL; // to store error code(s)
	HRESULT hr = pIOPCItemMgt->RemoveItems(1, hServerArray, &pErrors);
	//_ASSERT(!hr);

//	wcout <<"Method RemoveItem():"<<endl;
	if( (hr != S_OK) )
	{
		CurErr = GetErrorCSTRString(hr);
		DeleteLastSymbols(CurErr, 2);
		long lhr = HRESULT_CODE(hr);
		CurErrCode.Format(_T("%d"),lhr);
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
		//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);
	}
	else
	{
		AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);
	}

//	wcout <<ErrorString<<endl;
//	ErrorString.clear();


	//release memory allocated by the server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;
}


////////////////////////////////////////////////////////////////////////
// Remove the Group whose server handle is hServerGroup from the server
// whose IOPCServer interface is pointed by pIOPCServer
//
void CMFC_OPC_ExplorerDlg::RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	CurMethod = _T("RemoveGroup");
	CurParams = _T("hServerGroup");

	// Remove the group:
	HRESULT* pErrors=NULL; // to store error code(s)
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);

//	wcout <<"Method RemoveGroup():"<<endl;
	if( (hr != S_OK) )
	{
		CurErr = GetErrorCSTRString(hr);
		DeleteLastSymbols(CurErr, 2);
		long lhr = HRESULT_CODE(hr);
		CurErrCode.Format(_T("%d"),lhr);
		CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

		AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 2);

		LogSessionStatus(CurMethod, CurParamsAndErr, CurErrCode, true);
		//LogSessionStatus(CurMethod, CurParams, CurErrCode, false);
		//LogSessionStatus(CurMethod, CurErr, CurErrCode, true);
	}
	else
	{
		AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 2);
	}

//	wcout <<ErrorString<<endl;
//	ErrorString.clear();

	CoTaskMemFree(pErrors);
	pErrors = NULL;
	//_ASSERT(!hr);
}




int CMFC_OPC_ExplorerDlg::GetStatus(void)
{
	CurMethod = _T("GetStatus");
	CurParams = _T("");

	HRESULT hr_serv(NULL);
	OPCSERVERSTATUS *popcst(NULL);
	if(pIOPCServer)
	{
		hr_serv = pIOPCServer->GetStatus(&popcst);
		if(hr_serv==S_OK)
		{
		   if(popcst)
		   {
			int i=popcst->dwServerState;
			CoTaskMemFree(popcst);

			AddItemToListBox(_T("ОК - ") + CurMethod +_T(" | ")+ CurParams, 1);

			return (i);
		   }
		}
	}

	CurErr = GetErrorCSTRString(hr_serv);
	DeleteLastSymbols(CurErr, 2);
	long lhr = HRESULT_CODE(hr_serv);
	CurErrCode.Format(_T("%d"),lhr);
	CurParamsAndErr.Format(_T("%s | ОШИБКА: %s(%s)"), CurParams, CurErr, CurErrCode);

	AddItemToListBox(_T("ОШИБКА - ") + CurMethod +_T(" | ")+ CurParamsAndErr, 1);
	return 0;
}


//////////////////////////////////////////////////////////////////////////
// ПРОЧИЕ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
//////////////////////////////////////////////////////////////////////////

CString CMFC_OPC_ExplorerDlg::GetErrorCSTRString(HRESULT hrErr)
{
	CString str;
	HRESULT hr_scode=HRESULT_CODE(hrErr);
	HRESULT hr_facility=HRESULT_FACILITY(hrErr);
	HRESULT hr_severity=HRESULT_SEVERITY(hrErr);
//	if(hr_facility!=FACILITY_SYMBOL_OPC)
//	{
//		if(!pIOPCServer)
//		{
			LPVOID lpMsgBuf=0;
			if (!FormatMessage( 
				FORMAT_MESSAGE_ALLOCATE_BUFFER | 
				FORMAT_MESSAGE_FROM_SYSTEM | 
				FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL,
				hrErr,
				MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
				(LPTSTR) &lpMsgBuf,
				0,
				NULL ))
			{
				// Handle the error.
				return _T("Неизвестная ошибка  ");
			}
			// Process any inserts in lpMsgBuf.
			// ...
			// Display the string.
			if(lpMsgBuf)
			{
				str = (LPCTSTR)lpMsgBuf;
				LocalFree( lpMsgBuf );
			}
			else
				str=_T("Неизвестная ошибка  ");

			return str;
//		}


		LPWSTR pwStr=NULL;
		HRESULT hrT=MAKE_HRESULT(hr_severity,hr_facility,hr_scode);

	return str;
}




#define FACILITY_SYMBOL_OPC 72 

/*
CString CMFC_OPC_ExplorerDlg::GetErrorCSTRString(HRESULT hrErr)
{

	CString str;
	HRESULT hr_scode=HRESULT_CODE(hrErr);
	HRESULT hr_facility=HRESULT_FACILITY(hrErr);
	HRESULT hr_severity=HRESULT_SEVERITY(hrErr);
	if(hr_facility!=FACILITY_SYMBOL_OPC)
	{
		if(!pIOPCServer)
		{
			LPVOID lpMsgBuf=0;
			if (!FormatMessage( 
				FORMAT_MESSAGE_ALLOCATE_BUFFER | 
				FORMAT_MESSAGE_FROM_SYSTEM | 
				FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL,
				hrErr,
				MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
				(LPTSTR) &lpMsgBuf,
				0,
				NULL ))
			{
				// Handle the error.
				return _T("Неизвестная ошибка  ");
			}
			// Process any inserts in lpMsgBuf.
			// ...
			// Display the string.
			if(lpMsgBuf)
			{
				str = (LPCTSTR)lpMsgBuf;
				LocalFree( lpMsgBuf );
			}
			else
				str=_T("Неизвестная ошибка  ");

			return str;
		}


		LPWSTR pwStr=NULL;
		HRESULT hrT=MAKE_HRESULT(hr_severity,hr_facility,hr_scode);

		if(pIOPCServer->GetErrorString(hrErr,LOCALE_SYSTEM_DEFAULT,&pwStr)==S_OK)
		{
			str = pwStr;
			CoTaskMemFree(pwStr);

#if 0		
			char ch[512];
			WideCharToMultiByte( CP_ACP,0, pwStr, -1,
				ch, 512, NULL, NULL );
			//MultiByteToWideChar(CP_ACP,0, pwStr, -1,
			//	ch, 512, NULL, NULL );
			CoTaskMemFree(pwStr);

			std::string s = ch;
			vector<WCHAR> s2(100);// room for 100 characters
			//std::wstring final;
			//...
			MultiByteToWideChar(
			  CP_ACP,      // code page
			  0,         // character-type options
			  s.c_str(),         // string to map
			  s.length(),       // number of bytes in string
			  &s2[0],           // wide-character buffer (must use vector here!)
			  100                // size of buffer
			);
			str = &s2[0];  // Assign vector contents to true std::wstring.
#endif		

			//return	str;
        
		}
		else
		{
			//MessageBox(NULL,"косяк от сервера 1","Не удалось получить от ОРС сервера расшифровку ошибки",MB_OK);
			str=_T("Не удалось получить от ОРС сервера расшифровку ошибки  ");
		}
	};

	if(hr_facility==FACILITY_SYMBOL_OPC)
	{

	  //MessageBox(NULL,"FACILITY_SYMBOL_OPC","Не удалось получить от ОРС сервера расшифровку ошибки",MB_OK);
	  switch(hr_scode)
	  {
	  case E_ItemsPresent:			str= _T("Данный пункт присутствует в списке группы  ");break;
	  case E_Items_NO_Present:		str= _T("Данный пункт не присутствует в списке группы  ");break;
	  case E_Group_Present:			str= _T("Группа с указанным именем существует  ");break;
	  case E_Group_NO_Present:		str= _T("Группа с указанным именем не существует  ");break;
	  case E_Group_No_Adrres:		str= _T("Отсутсвует адрес группы  ");break;
	  case E_NO_Memory:				str= _T("Не удалось выделить требуемый размер памяти память  ");break;
	  case E_NO_AdrresIOPCAsyncIO2: str= _T("Отсутсвует адрес асинхронного интерфейса  ");break;
	  case E_NO_AdrresIOPCSyncIO:	str= _T("Отсутсвует адрес синхронного интерфейса  ");break;
	  case E_NO_AdrresIOPCItemMgt:	str= _T("Отсутсвует адрес  интерфейса значений  ");break;
	  case E_NO_IOPCServer:			str= _T("Не удается получить интерфес IOPCServer  ");break;
	  default :
		  str= _T("Неизветсная ошибка  ");
	  };
	}

	return str;
}
*/



std::wstring CMFC_OPC_ExplorerDlg::GetErrorSTDWString(HRESULT hrErr)
{
////	IOPCServer* pIOPCServer;   //pointer to IOPServer interface
////	IOPCItemMgt* pIOPCItemMgt;	//pointer to IOPCItemMgt interface

	std::wstring str;
	HRESULT hr_scode=HRESULT_CODE(hrErr);
	HRESULT hr_facility=HRESULT_FACILITY(hrErr);
	HRESULT hr_severity=HRESULT_SEVERITY(hrErr);
//	if(hr_facility<FACILITY_SYMBOL_OPC)
//	{
//		LPVOID lpMsgBuf;
//		if (!FormatMessage( 
//			FORMAT_MESSAGE_ALLOCATE_BUFFER | 
//			FORMAT_MESSAGE_FROM_SYSTEM | 
//			FORMAT_MESSAGE_IGNORE_INSERTS,
//			NULL,
//			hrErr,
//			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
//			(LPTSTR) &lpMsgBuf,
//			0,
//			NULL ))
//		{
//			// Handle the error.
//			return "Неизвестная ошибка";
//		}
		// Process any inserts in lpMsgBuf.
		// ...
		// Display the string.
//		str=(LPCTSTR)lpMsgBuf;
		// Free the buffer.
//		LocalFree( lpMsgBuf );
//		return str;
//	}
	if(hr_facility!=FACILITY_SYMBOL_OPC)
	{
		//hr_scode=HRESULT_CODE(hrErr);
		//hr_facility=HRESULT_FACILITY(hrErr);
		//hr_severity=HRESULT_SEVERITY(hrErr);
		//hr_facility=FACILITY_SYMBOL_OPC;
		
		//if(!pIOPServer)
		if(!pIOPCServer)
		{
			LPVOID lpMsgBuf=0;
			if (!FormatMessage( 
				FORMAT_MESSAGE_ALLOCATE_BUFFER | 
				FORMAT_MESSAGE_FROM_SYSTEM | 
				FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL,
				hrErr,
				MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
				(LPTSTR) &lpMsgBuf,
				0,
				NULL ))
			{
				// Handle the error.
				return _T("Неизвестная ошибка  ");
			}
			// Process any inserts in lpMsgBuf.
			// ...
			// Display the string.
			if(lpMsgBuf)
			{
				str = (LPCTSTR)lpMsgBuf;
				//str = (LPCSTR)lpMsgBuf;
				
				// Free the buffer.
				//if(lpMsgBuf)
				LocalFree( lpMsgBuf );
			}
			else
				str=_T("Неизвестная ошибка  ");

			return str;
		}


		//if(!pIOPServer)
		//	return string("Не удается получить интерфейс IOPCServer");
		LPWSTR pwStr=NULL;
		HRESULT hrT=MAKE_HRESULT(hr_severity,hr_facility,hr_scode);

//		if(pIOPServer->GetErrorString(hrErr,LOCALE_SYSTEM_DEFAULT,&pwStr)==S_OK)
//		if(pIOPCServer->GetErrorString(hrErr,LOCALE_SYSTEM_DEFAULT,&pwStr)==S_OK)
		if(pIOPCServer->GetErrorString(hrErr,LOCALE_SYSTEM_DEFAULT,&pwStr)==S_OK)
		{
			str = pwStr;
			CoTaskMemFree(pwStr);

			/*
			char ch[512];
			WideCharToMultiByte( CP_ACP,0, pwStr, -1,
				ch, 512, NULL, NULL );
			//MultiByteToWideChar(CP_ACP,0, pwStr, -1,
			//	ch, 512, NULL, NULL );
			CoTaskMemFree(pwStr);

			std::string s = ch;
			vector<WCHAR> s2(100);// room for 100 characters
			//std::wstring final;
			//...
			MultiByteToWideChar(
			  CP_ACP,      // code page
			  0,         // character-type options
			  s.c_str(),         // string to map
			  s.length(),       // number of bytes in string
			  &s2[0],           // wide-character buffer (must use vector here!)
			  100                // size of buffer
			);
			str = &s2[0];  // Assign vector contents to true std::wstring.
			*/

			//std::string chstr = ch;
			
			//str = StringToWString(chstr);
//			stringstream stre;
//			stre<<hex<<hrErr;
//			string ts;
//			stre>>ts;
//			MessageBox(NULL,ts.c_str(),str.c_str(),MB_OK);
			//return 
				str;
        
		}
		else
		{
			//MessageBox(NULL,"косяк от сервера 1","Не удалось получить от ОРС сервера расшифровку ошибки",MB_OK);
			str=_T("Не удалось получить от ОРС сервера расшифровку ошибки  ");
		}
	};

	if(hr_facility==FACILITY_SYMBOL_OPC)
	{

	  //MessageBox(NULL,"FACILITY_SYMBOL_OPC","Не удалось получить от ОРС сервера расшифровку ошибки",MB_OK);
	  switch(hr_scode)
	  {
	  case E_ItemsPresent:			str= _T("Данный пункт присутствует в списке группы  ");break;
	  case E_Items_NO_Present:		str= _T("Данный пункт не присутствует в списке группы  ");break;
	  case E_Group_Present:			str= _T("Группа с указанным именем существует  ");break;
	  case E_Group_NO_Present:		str= _T("Группа с указанным именем не существует  ");break;
	  case E_Group_No_Adrres:		str= _T("Отсутсвует адрес группы  ");break;
	  case E_NO_Memory:				str= _T("Не удалось выделить требуемый размер памяти память  ");break;
	  case E_NO_AdrresIOPCAsyncIO2: str= _T("Отсутсвует адрес асинхронного интерфейса  ");break;
	  case E_NO_AdrresIOPCSyncIO:	str= _T("Отсутсвует адрес синхронного интерфейса  ");break;
	  case E_NO_AdrresIOPCItemMgt:	str= _T("Отсутсвует адрес  интерфейса значений  ");break;
	  case E_NO_IOPCServer:			str= _T("Не удается получить интерфес IOPCServer  ");break;
	  default :
		  str= _T("Неизветсная ошибка  ");
	  };
	}

	return str;
}





///////////////////////////////////////////////////////////////////////////////////////
// МЕТОДЫ Конвертации /////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////

CString CMFC_OPC_ExplorerDlg::VariantTypeToCStrVarType(VARTYPE VariantTYPE)
{
	CString VariantTYPE_CStr("");

	switch (VariantTYPE)
	{
	case VT_EMPTY:				VariantTYPE_CStr = _T("VT_EMPTY");				break;
	case VT_NULL:				VariantTYPE_CStr = _T("VT_NULL");				break;
	case VT_BOOL:				VariantTYPE_CStr = _T("VT_BOOL");				break;
	case VT_I2:					VariantTYPE_CStr = _T("VT_I2");					break;
	case VT_I4:					VariantTYPE_CStr = _T("VT_I4");					break;
	case VT_R4:					VariantTYPE_CStr = _T("VT_R4");					break;
	case VT_R8:					VariantTYPE_CStr = _T("VT_R8");					break;
	case VT_BSTR:				VariantTYPE_CStr = _T("VT_BSTR");				break;
	case VT_DATE:				VariantTYPE_CStr = _T("VT_DATE");				break;
	case VT_VARIANT:			VariantTYPE_CStr = _T("VT_VARIANT");			break;
	case VT_CY:					VariantTYPE_CStr = _T("VT_CY");					break;
	case VT_RECORD:				VariantTYPE_CStr = _T("VT_RECORD");				break;
	case VT_LPSTR:				VariantTYPE_CStr = _T("VT_LPSTR");				break;
	case VT_LPWSTR:				VariantTYPE_CStr = _T("VT_LPWSTR");				break;
	case VT_ERROR:				VariantTYPE_CStr = _T("VT_ERROR");				break;
	case VT_I1:					VariantTYPE_CStr = _T("VT_I1");					break;
	case VT_UI1:				VariantTYPE_CStr = _T("VT_UI1");				break;
	case VT_UI2:				VariantTYPE_CStr = _T("VT_UI2");				break;
	case VT_UI4:				VariantTYPE_CStr = _T("VT_UI4");				break;
	case VT_I8:					VariantTYPE_CStr = _T("VT_I8");					break;
	case VT_UI8:				VariantTYPE_CStr = _T("VT_UI8");				break;
	case VT_INT:				VariantTYPE_CStr = _T("VT_INT");				break;
	case VT_UINT:				VariantTYPE_CStr = _T("VT_UINT");				break;
	case VT_INT_PTR:			VariantTYPE_CStr = _T("VT_INT_PTR");			break;
	case VT_UINT_PTR:			VariantTYPE_CStr = _T("VT_UINT_PTR");			break;
	case VT_HRESULT:			VariantTYPE_CStr = _T("VT_HRESULT");			break;
	case VT_VOID:				VariantTYPE_CStr = _T("VT_VOID");				break;
	case VT_DISPATCH: 			VariantTYPE_CStr = _T("VT_DISPATCH");			break;
	case VT_UNKNOWN: 			VariantTYPE_CStr = _T("VT_UNKNOWN");			break;
	case VT_DECIMAL: 			VariantTYPE_CStr = _T("VT_DECIMAL");			break;
	case VT_PTR:				VariantTYPE_CStr = _T("VT_PTR");				break;
	case VT_SAFEARRAY:			VariantTYPE_CStr = _T("VT_SAFEARRAY");			break;
	case VT_CARRAY:				VariantTYPE_CStr = _T("VT_CARRAY");				break;
	case VT_USERDEFINED:		VariantTYPE_CStr = _T("VT_USERDEFINED");		break;
	case VT_FILETIME:			VariantTYPE_CStr = _T("VT_FILETIME");			break;
	case VT_BLOB:				VariantTYPE_CStr = _T("VT_BLOB");				break;
	case VT_STREAM:				VariantTYPE_CStr = _T("VT_STREAM");				break;
	case VT_STORAGE:			VariantTYPE_CStr = _T("VT_STORAGE");			break;
	case VT_STREAMED_OBJECT:	VariantTYPE_CStr = _T("VT_STREAMED_OBJECT");	break;
	case VT_STORED_OBJECT:		VariantTYPE_CStr = _T("VT_STORED_OBJECT");		break;
	case VT_VERSIONED_STREAM:	VariantTYPE_CStr = _T("VT_VERSIONED_STREAM");	break;
	case VT_BLOB_OBJECT:		VariantTYPE_CStr = _T("VT_BLOB_OBJECT");		break;
	case VT_CF:					VariantTYPE_CStr = _T("VT_CF");					break;
	case VT_CLSID:				VariantTYPE_CStr = _T("VT_CLSID");				break;
	case VT_VECTOR:				VariantTYPE_CStr = _T("VT_VECTOR");				break;
	case VT_ARRAY:				VariantTYPE_CStr = _T("VT_ARRAY");				break;
	case VT_BYREF:				VariantTYPE_CStr = _T("VT_BYREF");				break;
	case VT_BSTR_BLOB:			VariantTYPE_CStr = _T("VT_BSTR_BLOB	");			break;

	default:
		break;
	}

	return VariantTYPE_CStr;
}


VARTYPE CMFC_OPC_ExplorerDlg::CStrVarTypeToVariantType(CString VariantTYPE_CStr)
{
	VARTYPE VariantTYPE;

	for(int i=0; i<1; i++)
	{//bool goout=false;
	 if( _T("VT_EMPTY")			== VariantTYPE_CStr)			 { VariantTYPE = VT_EMPTY;				break;}
	 else if( _T("VT_NULL")		== VariantTYPE_CStr)			 { VariantTYPE = VT_NULL;				break;}
	 else if( _T("VT_BOOL")		== VariantTYPE_CStr)			 { VariantTYPE = VT_BOOL;				break;}
	 else if( _T("VT_I2")		== VariantTYPE_CStr)			 { VariantTYPE = VT_I2;					break;}
	 else if( _T("VT_I4")		== VariantTYPE_CStr)			 { VariantTYPE = VT_I4;					break;}
	 else if( _T("VT_R4")		== VariantTYPE_CStr)			 { VariantTYPE = VT_R4;					break;}
	 else if( _T("VT_R8")		== VariantTYPE_CStr)			 { VariantTYPE = VT_R8;					break;}
	 else if( _T("VT_BSTR")		== VariantTYPE_CStr)			 { VariantTYPE = VT_BSTR;				break;}
	 else if( _T("VT_LPSTR")	== VariantTYPE_CStr)			 { VariantTYPE = VT_LPSTR;				break;}
	 else if( _T("VT_LPWSTR")	== VariantTYPE_CStr)			 { VariantTYPE = VT_LPWSTR;				break;}
	 else if( _T("VT_DATE")		== VariantTYPE_CStr)			 { VariantTYPE = VT_DATE;				break;}
	 else if( _T("VT_ERROR")	== VariantTYPE_CStr)			 { VariantTYPE = VT_ERROR;				break;}
	 else if( _T("VT_VARIANT")	== VariantTYPE_CStr)			 { VariantTYPE = VT_VARIANT;			break;}
	 else if( _T("VT_I1")		== VariantTYPE_CStr)			 { VariantTYPE = VT_I1;					break;}
	 else if( _T("VT_UI1")		== VariantTYPE_CStr)			 { VariantTYPE = VT_UI1;				break;}
	 else if( _T("VT_UI2")		== VariantTYPE_CStr)			 { VariantTYPE = VT_UI2;				break;}
	 else if( _T("VT_UI4")		== VariantTYPE_CStr)			 { VariantTYPE = VT_UI4;				break;}
	 else if( _T("VT_I8")		== VariantTYPE_CStr)			 { VariantTYPE = VT_I8;					break;}
	 else if( _T("VT_UI8")		== VariantTYPE_CStr)			 { VariantTYPE = VT_UI8;				break;}
	 else if( _T("VT_INT")		== VariantTYPE_CStr)			 { VariantTYPE = VT_INT;				break;}
	 else if( _T("VT_UINT")		== VariantTYPE_CStr)			 { VariantTYPE = VT_UINT;				break;}
	 else if( _T("VT_INT_PTR")	== VariantTYPE_CStr)			 { VariantTYPE = VT_INT_PTR;			break;}
	 else if( _T("VT_UINT_PTR")	== VariantTYPE_CStr)			 { VariantTYPE = VT_UINT_PTR;			break;}
	 else if( _T("VT_DECIMAL")	== VariantTYPE_CStr)			 { VariantTYPE = VT_DECIMAL;			break;}
	 else if( _T("VT_HRESULT")	== VariantTYPE_CStr)			 { VariantTYPE = VT_HRESULT;			break;}
	 else if( _T("VT_VOID")		== VariantTYPE_CStr)			 { VariantTYPE = VT_VOID;				break;}
	 else if( _T("VT_CY")		== VariantTYPE_CStr)			 { VariantTYPE = VT_CY;					break;}
	 else if( _T("VT_DISPATCH")	== VariantTYPE_CStr)			 { VariantTYPE = VT_DISPATCH;			break;}
	 else if( _T("VT_UNKNOWN")	== VariantTYPE_CStr)			 { VariantTYPE = VT_UNKNOWN;			break;}
	 else if( _T("VT_RECORD")	== VariantTYPE_CStr)			 { VariantTYPE = VT_RECORD;				break;}
	 else if( _T("VT_PTR")		== VariantTYPE_CStr)			 { VariantTYPE = VT_PTR;				break;}
	 else if( _T("VT_SAFEARRAY")== VariantTYPE_CStr)			 { VariantTYPE = VT_SAFEARRAY;			break;}
	 else if( _T("VT_CARRAY")	== VariantTYPE_CStr)			 { VariantTYPE = VT_CARRAY;				break;}
	 else if(_T("VT_USERDEFINED")==VariantTYPE_CStr)			 { VariantTYPE = VT_USERDEFINED;		break;}
	 else if( _T("VT_FILETIME")	== VariantTYPE_CStr)			 { VariantTYPE = VT_FILETIME;			break;}
	 else if( _T("VT_BLOB")		== VariantTYPE_CStr)			 { VariantTYPE = VT_BLOB;				break;}
	 else if( _T("VT_STREAM")	== VariantTYPE_CStr)			 { VariantTYPE = VT_STREAM;				break;}
	 else if( _T("VT_STORAGE")	== VariantTYPE_CStr)			 { VariantTYPE = VT_STORAGE;			break;}
	 else if( _T("VT_STREAMED_OBJECT")	== VariantTYPE_CStr)	 { VariantTYPE = VT_STREAMED_OBJECT;	break;}
	 else if( _T("VT_STORED_OBJECT")	== VariantTYPE_CStr)	 { VariantTYPE = VT_STORED_OBJECT;		break;}
	 else if( _T("VT_VERSIONED_STREAM")	== VariantTYPE_CStr)	 { VariantTYPE = VT_VERSIONED_STREAM;	break;}
	 else if( _T("VT_BLOB_OBJECT")		== VariantTYPE_CStr)	 { VariantTYPE = VT_BLOB_OBJECT;		break;}
	 else if( _T("VT_CF")		== VariantTYPE_CStr)			 { VariantTYPE = VT_CF;					break;}
	 else if( _T("VT_CLSID")	== VariantTYPE_CStr)			 { VariantTYPE = VT_CLSID;				break;}
	 else if( _T("VT_VECTOR")	== VariantTYPE_CStr)			 { VariantTYPE = VT_VECTOR;				break;}
	 else if( _T("VT_ARRAY")	== VariantTYPE_CStr)			 { VariantTYPE = VT_ARRAY;				break;}
	 else if( _T("VT_BYREF")	== VariantTYPE_CStr)			 { VariantTYPE = VT_BYREF;				break;}
	 else if( _T("VT_BSTR_BLOB")== VariantTYPE_CStr)			 { VariantTYPE = VT_BSTR_BLOB;			break;}
	 else { VariantTYPE = 1234567890; }
	 //if(!goout) {break;}
	}

	return VariantTYPE;
}


//VARINAT COPCTestDlg::CStringToVariant(CString VariantVALUE_CStr, CString VariantTYPE_CStr)
_variant_t CMFC_OPC_ExplorerDlg::CStrToVariant(CString VariantVALUE_CStr, CString VariantTYPE_CStr)
{
	VARIANT VariantVARIANT;


	VARTYPE VariantTYPE;
	VariantTYPE = CStrVarTypeToVariantType(VariantTYPE_CStr);

//////////////////////////////////////////////////////////
// Подготавливаем Variant переменную для записи и отправки

	//VARIANT varVal;	//_variant_t varVal;
	VariantVARIANT.vt = VariantTYPE;

	switch(VariantVARIANT.vt)
	{
	case VT_EMPTY: break;
	case VT_NULL: break;
	case VT_BOOL:
		if(CStrToBool(VariantVALUE_CStr))
		{VariantVARIANT.boolVal=true;}
		if( ! CStrToBool(VariantVALUE_CStr) )
		{VariantVARIANT.boolVal=false;}
		break;
	case VT_I2:
		VariantVARIANT.iVal = _wtoi(VariantVALUE_CStr);
		break;
	case VT_I4:
		VariantVARIANT.lVal = _wtol(VariantVALUE_CStr);
		break;
	case VT_R4:
		VariantVARIANT.fltVal = (float)_wtof(VariantVALUE_CStr);
		break;
	case VT_R8:
		VariantVARIANT.dblVal = _wtof(VariantVALUE_CStr);
		break;
	case VT_BSTR:
		VariantVARIANT.bstrVal = VariantVALUE_CStr.AllocSysString();
		break;
/*
	case VT_LPSTR: break;
	case VT_LPWSTR: break;
	case VT_DATE: break;
		CStr_Res = BoolToCStr(VariantVARIANT.date);
		break;
	case VT_ERROR: break;
	case VT_VARIANT:
//		CStr_Res = (VariantVARIANT.vt);
		break;
	case VT_I1: break;
	case VT_UI1: break;
	case VT_UI2: break;
	case VT_UI4: break;
	case VT_I8: break;
	case VT_UI8: break;
	case VT_INT: break;
	case VT_UINT: break;
	case VT_INT_PTR: break;
	case VT_UINT_PTR: break;
	case VT_HRESULT: break;
	case VT_VOID: break;
	case VT_CY:
//		CStr_Res = (VariantVARIANT.cyVal);
		break;
	case VT_DISPATCH: break;
	case VT_UNKNOWN: break;
	case VT_DECIMAL: break;
	case VT_RECORD:
//		CStr_Res = (VariantVARIANT.pvRecord);
		break;
	case VT_PTR: break;
	case VT_SAFEARRAY: break;
	case VT_CARRAY: break;
	case VT_USERDEFINED: break;
	case VT_FILETIME: break;
	case VT_BLOB: break;
	case VT_STREAM: break;
	case VT_STORAGE: break;
	case VT_STREAMED_OBJECT: break;
	case VT_STORED_OBJECT: break;
	case VT_VERSIONED_STREAM: break;
	case VT_BLOB_OBJECT: break;
	case VT_CF: break;
	case VT_CLSID: break;
	case VT_VECTOR: break;
	case VT_ARRAY: break;
	case VT_BYREF: break;
	case VT_BSTR_BLOB: break;
*/
	default:
		VariantVARIANT.vt = 1234567890;
		//AfxMessageBox(_T("Данный тип данных необходимо внедрить в функцию обработки!"));
		break;
	}

	return VariantVARIANT;
}



#if 0
/*
 * VARENUM usage key,
 *
 * * [V] - may appear in a VARIANT
 * * [T] - may appear in a TYPEDESC
 * * [P] - may appear in an OLE property set
 * * [S] - may appear in a Safe Array
 *
 *
 *  VT_EMPTY            [V]   [P]     nothing
 *  VT_NULL             [V]   [P]     SQL style Null
 *  VT_I2               [V][T][P][S]  2 byte signed int
 *  VT_I4               [V][T][P][S]  4 byte signed int
 *  VT_R4               [V][T][P][S]  4 byte real
 *  VT_R8               [V][T][P][S]  8 byte real
 *  VT_CY               [V][T][P][S]  currency
 *  VT_DATE             [V][T][P][S]  date
 *  VT_BSTR             [V][T][P][S]  OLE Automation string
 *  VT_DISPATCH         [V][T]   [S]  IDispatch *
 *  VT_ERROR            [V][T][P][S]  SCODE
 *  VT_BOOL             [V][T][P][S]  True=-1, False=0
 *  VT_VARIANT          [V][T][P][S]  VARIANT *
 *  VT_UNKNOWN          [V][T]   [S]  IUnknown *
 *  VT_DECIMAL          [V][T]   [S]  16 byte fixed point
 *  VT_RECORD           [V]   [P][S]  user defined type
 *  VT_I1               [V][T][P][s]  signed char
 *  VT_UI1              [V][T][P][S]  unsigned char
 *  VT_UI2              [V][T][P][S]  unsigned short
 *  VT_UI4              [V][T][P][S]  unsigned long
 *  VT_I8                  [T][P]     signed 64-bit int
 *  VT_UI8                 [T][P]     unsigned 64-bit int
 *  VT_INT              [V][T][P][S]  signed machine int
 *  VT_UINT             [V][T]   [S]  unsigned machine int
 *  VT_INT_PTR             [T]        signed machine register size width
 *  VT_UINT_PTR            [T]        unsigned machine register size width
 *  VT_VOID                [T]        C style void
 *  VT_HRESULT             [T]        Standard return type
 *  VT_PTR                 [T]        pointer type
 *  VT_SAFEARRAY           [T]        (use VT_ARRAY in VARIANT)
 *  VT_CARRAY              [T]        C style array
 *  VT_USERDEFINED         [T]        user defined type
 *  VT_LPSTR               [T][P]     null terminated string
 *  VT_LPWSTR              [T][P]     wide null terminated string
 *  VT_FILETIME               [P]     FILETIME
 *  VT_BLOB                   [P]     Length prefixed bytes
 *  VT_STREAM                 [P]     Name of the stream follows
 *  VT_STORAGE                [P]     Name of the storage follows
 *  VT_STREAMED_OBJECT        [P]     Stream contains an object
 *  VT_STORED_OBJECT          [P]     Storage contains an object
 *  VT_VERSIONED_STREAM       [P]     Stream with a GUID version
 *  VT_BLOB_OBJECT            [P]     Blob contains an object 
 *  VT_CF                     [P]     Clipboard format
 *  VT_CLSID                  [P]     A Class ID
 *  VT_VECTOR                 [P]     simple counted array
 *  VT_ARRAY            [V]           SAFEARRAY*
 *  VT_BYREF            [V]           void* for local use
 *  VT_BSTR_BLOB                      Reserved for system use
 */

enum VARENUM
    {	VT_EMPTY	= 0,
	VT_NULL	= 1,
	VT_I2	= 2,
	VT_I4	= 3,
	VT_R4	= 4,
	VT_R8	= 5,
	VT_CY	= 6,
	VT_DATE	= 7,
	VT_BSTR	= 8,
	VT_DISPATCH	= 9,
	VT_ERROR	= 10,
	VT_BOOL	= 11,
	VT_VARIANT	= 12,
	VT_UNKNOWN	= 13,
	VT_DECIMAL	= 14,
	VT_I1	= 16,
	VT_UI1	= 17,
	VT_UI2	= 18,
	VT_UI4	= 19,
	VT_I8	= 20,
	VT_UI8	= 21,
	VT_INT	= 22,
	VT_UINT	= 23,
	VT_VOID	= 24,
	VT_HRESULT	= 25,
	VT_PTR	= 26,
	VT_SAFEARRAY	= 27,
	VT_CARRAY	= 28,
	VT_USERDEFINED	= 29,
	VT_LPSTR	= 30,
	VT_LPWSTR	= 31,
	VT_RECORD	= 36,
	VT_INT_PTR	= 37,
	VT_UINT_PTR	= 38,
	VT_FILETIME	= 64,
	VT_BLOB	= 65,
	VT_STREAM	= 66,
	VT_STORAGE	= 67,
	VT_STREAMED_OBJECT	= 68,
	VT_STORED_OBJECT	= 69,
	VT_BLOB_OBJECT	= 70,
	VT_CF	= 71,
	VT_CLSID	= 72,
	VT_VERSIONED_STREAM	= 73,
	VT_BSTR_BLOB	= 0xfff,
	VT_VECTOR	= 0x1000,
	VT_ARRAY	= 0x2000,
	VT_BYREF	= 0x4000,
	VT_RESERVED	= 0x8000,
	VT_ILLEGAL	= 0xffff,
	VT_ILLEGALMASKED	= 0xfff,
	VT_TYPEMASK	= 0xfff
    } ;
#endif


CString CMFC_OPC_ExplorerDlg::VariantValueToCStr(VARIANT varVal)
{
	CString CStr_Res("");
	CString sDateFormat		= _T("%d/%m/%y");

	CString sNotAddedType	= _T("Данный тип данных ещё не внедрён в функцию обработки!");
//	CString sNotAddedType	= _T("НЕТ_ТИПА");
	CString sNoType			= _T("Отсуствует указанный тип или такой тип данных не поддерживается либо было разорвано соединение с сервером!");

	switch(varVal.vt)
	{
	case VT_EMPTY: break;	//=0	(Empty) No value
	case VT_NULL: break;	//=1	(Null) null.
	case VT_I2:					CStr_Res = LUintToCStr(varVal.iVal);	break;	//=2	(Short) two-byte integer.
	case VT_I4:					CStr_Res = LUintToCStr(varVal.lVal);	break;	//=3	(Int) four-byte integer.
	case VT_R4:					CStr_Res = FloatToCStr(varVal.fltVal);	break;	//=4	(Float) single-precision (four-byte) floating point number.
	case VT_R8:					CStr_Res = DoubleToCStr(varVal.dblVal);	break;	//=5	(Double) double-precision (8-byte) floating point number.
	case VT_CY:					CStr_Res = sNotAddedType;				break;	//=6	//CStr_Res = (varVal.cyVal)  //(Currency) currency amount.
	case VT_DATE:				CStr_Res = FormatOleDateTimeToCStr(
											varVal.date, sDateFormat);	break;	//=7	(Date) date.
	case VT_BSTR:				CStr_Res = varVal.bstrVal;				break;	//=8	(String) string.
	case VT_DISPATCH: 			CStr_Res = sNotAddedType;				break;	//=9  [NO] (Dispatch) dispatch interface.
	case VT_ERROR:				CStr_Res = sNotAddedType;				break;	//=10 [NO] (Error) error value.
	case VT_BOOL:				CStr_Res = BoolToCStr(varVal.boolVal);	break;	//=11	(Boolean) Boolean.
//*
	case VT_VARIANT:			CStr_Res = sNotAddedType;				break;	//=12 [NO] //CStr_Res = (varVal.vt)  //(Variant) variant.
	case VT_UNKNOWN: 			CStr_Res = sNotAddedType;				break;	//=13 [NO] (Object) object.
	case VT_DECIMAL: 			CStr_Res = sNotAddedType;				break;	//=14 [NO]
	case VT_I1:					CStr_Res = LUintToCStr(varVal.iVal);	break;	//=16 [NO] //???char
	case VT_UI1:				CStr_Res = LUintToCStr(varVal.iVal);	break;	//=17 [NO] (Byte) byte.
	case VT_UI2:				CStr_Res = LUintToCStr(varVal.iVal);	break;	//=18 [NO]
	case VT_UI4:				CStr_Res = LUintToCStr(varVal.iVal);	break;	//=19 [NO]
	case VT_I8:					CStr_Res = LUintToCStr(varVal.iVal);	break;	//=20 [NO]
	case VT_UI8:				CStr_Res = LUintToCStr(varVal.iVal);	break;	//=21 [NO]
	case VT_INT:				CStr_Res = LUintToCStr(varVal.iVal);	break;	//=22 [NO]
	case VT_UINT:				CStr_Res = LUintToCStr(varVal.iVal);	break;	//=23 [NO]
	case VT_VOID:				CStr_Res = sNotAddedType;				break;	//=24 [NO]
	case VT_HRESULT:			CStr_Res = sNotAddedType;				break;	//=25 [NO]
	case VT_PTR:				CStr_Res = sNotAddedType;				break;	//=26 [NO]
	case VT_SAFEARRAY:			CStr_Res = sNotAddedType;				break;	//=27 [NO]
	case VT_CARRAY:				CStr_Res = sNotAddedType;				break;	//=28 [NO]
	case VT_USERDEFINED:		CStr_Res = sNotAddedType;				break;	//=29 [NO]
	case VT_LPSTR:				CStr_Res = sNotAddedType;				break;	//=30 [NO] 
	case VT_LPWSTR:				CStr_Res = sNotAddedType;				break;	//=31 [NO]
	case VT_RECORD:				CStr_Res = sNotAddedType;				break;	//=36 [NO] //CStr_Res = (varVal.pvRecord);
	case VT_INT_PTR:			CStr_Res = sNotAddedType;				break;	//=37 [NO]
	case VT_UINT_PTR:			CStr_Res = sNotAddedType;				break;	//=38 [NO]
	case VT_FILETIME:			CStr_Res = sNotAddedType;				break;	//=64 [NO]
	case VT_BLOB:				CStr_Res = sNotAddedType;				break;	//=65 [NO]
	case VT_STREAM:				CStr_Res = sNotAddedType;				break;	//=66 [NO]
	case VT_STORAGE:			CStr_Res = sNotAddedType;				break;	//=67 [NO]
	case VT_STREAMED_OBJECT:	CStr_Res = sNotAddedType;				break;	//=68 [NO]
	case VT_STORED_OBJECT:		CStr_Res = sNotAddedType;				break;	//=69 [NO]
	case VT_BLOB_OBJECT:		CStr_Res = sNotAddedType;				break;	//=70 [NO]
	case VT_CF:					CStr_Res = sNotAddedType;				break;	//=71 [NO]
	case VT_CLSID:				CStr_Res = sNotAddedType;				break;	//=72 [NO]
	case VT_VERSIONED_STREAM:	CStr_Res = sNotAddedType;				break;	//=73 [NO]
	//case VT_BSTR_BLOB:		CStr_Res = sNotAddedType;				break;	//=4095  (0xfff)  [NO] 
	//case VT_ILLEGALMASKED:	CStr_Res = sNotAddedType;				break;	//=4095  (0xfff)  [NO] 
	case VT_TYPEMASK:			CStr_Res = sNotAddedType;				break;	//=4095  (0xfff)  [NO] (VariantTypeMask) Useful if you don't care about the type of the value being passed, but are interested in the "shape," that is, whether it is a VariantByref or not.
	case VT_VECTOR:				CStr_Res = sNotAddedType;				break;	//=4096	 (0x1000) [NO]
	case VT_ARRAY:				CStr_Res = sNotAddedType;				break;	//=8192  (0x2000) [NO] (VariantArray) Array of values of the given type.
	case VT_BYREF:				CStr_Res = sNotAddedType;				break;	//=16384 (0x4000) [NO]  //CStr_Res = (varVal.byref) // (VariantByref) Value is being passed by reference; any changes made to the parameter value by the called party are visible to the caller.
	case VT_RESERVED:			CStr_Res = sNotAddedType;				break;	//=32768 (0x8000) [NO]  //CStr_Res = (varVal.byref) // (VariantByref) Value is being passed by reference; any changes made to the parameter value by the called party are visible to the caller.
	case VT_ILLEGAL:			CStr_Res = sNotAddedType;				break;	//=65535 (0xffff) [NO]  //CStr_Res = (varVal.byref) // (VariantByref) Value is being passed by reference; any changes made to the parameter value by the called party are visible to the caller.
//*/

	default:
		CStr_Res = sNoType;
		break;
	}


	return CStr_Res;
}


/*
#define    OPC_QUALITY_MASK            0xC0
#define    OPC_STATUS_MASK             0xFC
#define    OPC_LIMIT_MASK              0x03
#define    OPC_QUALITY_BAD             0x00
#define    OPC_QUALITY_UNCERTAIN       0x40
#define    OPC_QUALITY_GOOD            0xC0
#define    OPC_QUALITY_CONFIG_ERROR    0x04
#define    OPC_QUALITY_NOT_CONNECTED   0x08
#define    OPC_QUALITY_DEVICE_FAILURE  0x0c
#define    OPC_QUALITY_SENSOR_FAILURE  0x10
#define    OPC_QUALITY_LAST_KNOWN      0x14
#define    OPC_QUALITY_COMM_FAILURE    0x18
#define    OPC_QUALITY_OUT_OF_SERVICE  0x1C
#define    OPC_QUALITY_LAST_USABLE     0x44
#define    OPC_QUALITY_SENSOR_CAL      0x50
#define    OPC_QUALITY_EGU_EXCEEDED    0x54
#define    OPC_QUALITY_SUB_NORMAL      0x58
#define    OPC_QUALITY_LOCAL_OVERRIDE  0xD8
#define    OPC_LIMIT_OK                0x00
#define    OPC_LIMIT_LOW               0x01
#define    OPC_LIMIT_HIGH              0x02
#define    OPC_LIMIT_CONST             0x03
*/





NumericindicarorLook CMFC_OPC_ExplorerDlg::GetTypeOfQual(WORD wrdValue)
{
	if(OPC_QUALITY_GOOD == wrdValue)
	{
		//AfxMessageBox(_T("GOOD"));
		return TAG_GOOD;
	}
	else
	{
		switch(wrdValue&OPC_STATUS_MASK)
		{
			// Н/Д
			case OPC_QUALITY_BAD:			//AfxMessageBox(_T("BAD NON SPECIFIC"));
			return BAD_ND;
			case OPC_QUALITY_LAST_KNOWN:	//AfxMessageBox(_T("BAD LAST KNOWN"));
			return BAD_ND;


			case OPC_QUALITY_UNCERTAIN:		//AfxMessageBox(_T("BAD NON UNCERTAIN"));
			switch(wrdValue&OPC_LIMIT_MASK)
			{
			case OPC_LIMIT_LOW:	 // АВ
			{return BAD_AV;}
			case OPC_LIMIT_HIGH: // КЗ
			{return BAD_KZ;}
			}
		}
		return BAD_OTHER;
	}
	return TAG_ERROR;
}

/*
//Качество значения для дискретных сигналов:
OverlayPictureType CMFC_OPC_ExplorerDlg::GetSiteSignalPicType( ItemValue& val )
{
	OverlayPictureType enPicType = EMPTY;

	if(V_VT(&val.Value)==VT_BOOL)
	{
		switch(val.Quality&OPC_QUALITY_MASK)
		{
		case OPC_QUALITY_GOOD:
			enPicType = (val.Quality&OPC_QUALITY_DEVICE_TIMEOUT)? 
				(V_BOOL(&val.Value)==VARIANT_TRUE? NONSIGNALED_TO : SIGNALED_TO)
				:
				(V_BOOL(&val.Value)==VARIANT_TRUE? NONSIGNALED    : SIGNALED   );
			break;

		case OPC_QUALITY_BAD:
			switch(val.Quality&OPC_STATUS_MASK)
			{
			case OPC_QUALITY_LAST_KNOWN:
				enPicType = (val.Quality&OPC_QUALITY_DEVICE_TIMEOUT)? 
					(V_BOOL(&val.Value)==VARIANT_TRUE? NONSIGNALED_TO : SIGNALED_TO)
					:
					(V_BOOL(&val.Value)==VARIANT_TRUE? NONSIGNALED    : SIGNALED   );
				break;
			case OPC_QUALITY_SENSOR_FAILURE://"undef"
				enPicType = (val.Quality&OPC_QUALITY_DEVICE_TIMEOUT)? UNDEF_TO : UNDEF;
				break;
			case OPC_QUALITY_COMM_FAILURE://"error"
				enPicType = (val.Quality&OPC_QUALITY_DEVICE_TIMEOUT)? BREAK_TO : BREAK;
				break;
			default:
				enPicType = TIMEOUT;
			}
			break;
		}
	}

	return enPicType;
}
*/


CString CMFC_OPC_ExplorerDlg::OneTagPropsToFullCStr(IOPCServer* pIOPCServer, TagStruct OneTag)
{
	CString sTagProperties("");

	try{
		HRESULT hr;

		//Адресное пространство сервера
		//запросить интерфейс
		IOPCItemProperties* pIProps = NULL;
		hr = pIOPCServer->QueryInterface(IID_IOPCItemProperties,(void**)&pIProps);
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer: failed QueryInterface(IOPCItemProperties): "
			//		<<hres(hr)<<AppDiag.End();
			throw hr;
		}

		//запросить список значений свойств тэга
		//# define PROPS_NUM 6
		//DWORD adwProps[PROPS_NUM] = { 100, 101, 5000, 5001, 5002, 5003 };

		# define PROPS_NUM 4
		DWORD adwProps[PROPS_NUM] = { 10000, 10001, 10002, 10006 };

		/*DWORD adwProps[11] = {
		1 - Canonical Data Type	- Type
		2 - Item Value			- Val
		3 - Item Quality		- Short Interger
		4 - Item Timestamp		- Date
		5 - Item Access Rights	- Long Integer
		6 - Server Scan Rate	- Single Float
		100 - EU Units			- String
		101 - Item Description	- String
		5000 - Associated Product Name - String
		5001 - Physical quantity - String
		5002 - Site Description	- String
		5003 - Site ID			- Unsigned Long
		};
		*/

		VARIANT* pPropValues = NULL;
		HRESULT* pPropValuesErrors = NULL;

		USES_CONVERSION;		// Cstring to LPWSTR
		//LPWSTR psz = CT2W(s);
		CString sTagFullName = OneTag.TagFullName;
		LPWSTR psz = (LPWSTR)(LPCTSTR)sTagFullName;

		hr = pIProps->GetItemProperties(
										//(LPWSTR)strTag.c_str()
										psz
										, PROPS_NUM
										, adwProps
										, &pPropValues
										, &pPropValuesErrors
									  );
		if(FAILED(hr)) {
			//AppDiag.Record(2)<<"COPCServer::GetItemProperties failed: "
			//		<<hres(hr)<<AppDiag.End();
			pIProps->Release();
			throw hr;
		}



		for(DWORD dwI=0; dwI<PROPS_NUM; dwI++)
		{
			if(pPropValuesErrors[dwI]==S_OK)
			{
				//Props[ adwProps[dwI] ] = pPropValues[dwI];
				//v_Props.push_back( VariantToCStr(pPropValues[dwI]) );

				sTagProperties = sTagProperties + ( LUintToCStr(adwProps[dwI]) ) + _T(": ");
				sTagProperties = sTagProperties + ( VariantValueToCStr(pPropValues[dwI]) ) + _T("\n");
			}
		}

		//освобождаем выделенную сервером память
		::CoTaskMemFree(pPropValues);
		::CoTaskMemFree(pPropValuesErrors);

		//освобождаем интерфейс
		pIProps->Release();
	}
	catch(HRESULT hError)
	{
		hError;

//		return false;
	}


	if(sTagProperties.IsEmpty()) sTagProperties = _T("PROPERTIES_ERROR");


	return sTagProperties;
}


CString CMFC_OPC_ExplorerDlg::QualToFullCStr(WORD wrdValue)
{
	CString sQuality("");
	sQuality = _T("OUR_QUALITY");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue) +_T("]") +_T("\n");

	switch(wrdValue & OPC_QUALITY_MASK)
	{
	case OPC_QUALITY_GOOD:					sQuality = _T("OPC_QUALITY_GOOD");	sQuality=sQuality +_T(" [")+ LUintToCStr(OPC_QUALITY_GOOD) +_T("]");	break;
		
	case OPC_QUALITY_BAD:					sQuality = _T("OPC_QUALITY_BAD");	sQuality=sQuality +_T(" [")+ LUintToCStr(OPC_QUALITY_BAD) +_T("]") +_T("\n");
		switch(wrdValue & OPC_STATUS_MASK)
		{
		case OPC_QUALITY_LAST_KNOWN:		sQuality += _T("OPC_QUALITY_LAST_KNOWN");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_CONFIG_ERROR:		sQuality += _T("OPC_QUALITY_CONFIG_ERROR");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_NOT_CONNECTED:		sQuality += _T("OPC_QUALITY_NOT_CONNECTED");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_DEVICE_FAILURE:	sQuality += _T("OPC_QUALITY_DEVICE_FAILURE");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_SENSOR_FAILURE:	sQuality += _T("OPC_QUALITY_SENSOR_FAILURE");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;	//"undef"
		case OPC_QUALITY_COMM_FAILURE:		sQuality += _T("OPC_QUALITY_COMM_FAILURE");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;	//"error"
		case OPC_QUALITY_OUT_OF_SERVICE:	sQuality += _T("OPC_QUALITY_OUT_OF_SERVICE");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_LAST_USABLE:		sQuality += _T("OPC_QUALITY_LAST_USABLE");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_SENSOR_CAL:		sQuality += _T("OPC_QUALITY_SENSOR_CAL");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_EGU_EXCEEDED:		sQuality += _T("OPC_QUALITY_EGU_EXCEEDED");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_SUB_NORMAL:		sQuality += _T("OPC_QUALITY_SUB_NORMAL");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		case OPC_QUALITY_LOCAL_OVERRIDE:	sQuality += _T("OPC_QUALITY_LOCAL_OVERRIDE");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_STATUS_MASK) +_T("]");	break;
		}
		break;

	case OPC_QUALITY_UNCERTAIN:				sQuality += _T("OPC_QUALITY_UNCERTAIN");		sQuality=sQuality +_T(" [")+ LUintToCStr(OPC_QUALITY_UNCERTAIN) +_T("]") +_T("\n");
		switch(wrdValue & OPC_LIMIT_MASK)
		{
		case OPC_LIMIT_OK:					sQuality += _T("OPC_LIMIT_OK");		sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_LIMIT_MASK) +_T("]");	break;
		case OPC_LIMIT_LOW:					sQuality += _T("OPC_LIMIT_LOW");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_LIMIT_MASK) +_T("]");	break;	// АВ
		case OPC_LIMIT_HIGH:				sQuality += _T("OPC_LIMIT_HIGH");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_LIMIT_MASK) +_T("]");	break;	// КЗ
		case OPC_LIMIT_CONST:				sQuality += _T("OPC_LIMIT_CONST");	sQuality=sQuality +_T(" [")+ LUintToCStr(wrdValue & OPC_LIMIT_MASK) +_T("]");	break;	// КЗ
		}
		break;
		//default:
			//TIMEOUT;
	}

	if(sQuality.IsEmpty())	sQuality = _T("UNKNOWN_QUALITY");


	return sQuality;
}


CString CMFC_OPC_ExplorerDlg::QualToBoolCStr(bool w, bool BigFirstLetter)
{
	if(BigFirstLetter)	return( w ? _T("Good") : _T("Bad") );
	else				return( w ? _T("good") : _T("bad") );
}


bool CMFC_OPC_ExplorerDlg::CStrToBool(CString cstr)
{
	if(cstr == _T("true")  || cstr == _T("True"))	{return true;}
	if(cstr == _T("false") || cstr == _T("False"))	{return false;}
}

CString CMFC_OPC_ExplorerDlg::BoolToCStr(bool b, bool BigFirstLetter)
{
	if(BigFirstLetter)	return( b ? _T("True") : _T("False") );
	else				return( b ? _T("true") : _T("false") );
}

CString CMFC_OPC_ExplorerDlg::LUintToCStr(__int32 Num)
{
	CString CStr;
	CStr.Format( _T("%I32u"), Num ); //%i int //%u uint //%l long //%ll longlong // %I64u Ulonglong //%5d int-right5
	return CStr;
}
CString CMFC_OPC_ExplorerDlg::LUintToCStr(unsigned int Num)
{
	CString CStr;
	CStr.Format( _T("%u"), Num ); //%i int //%u uint //%l long //%ll longlong // %I64u Ulonglong //%5d int-right5
	return CStr;
}
CString CMFC_OPC_ExplorerDlg::LUintToCStr(long Num)
{
	CString CStr;
	CStr.Format( _T("%l"), Num ); //%i int //%u uint //%l long //%ll longlong // %I64u Ulonglong //%5d int-right5
	return CStr;
}
CString CMFC_OPC_ExplorerDlg::LUintToCStr(unsigned long Num)
{
	CString CStr;
	CStr.Format( _T("%lu"), Num ); //%i int //%u uint //%l long //%lu ulong //%ll longlong // %I64u Ulonglong //%5d int-right5
	return CStr;
}
CString CMFC_OPC_ExplorerDlg::LUintToCStr(__int64 Num)
{
	CString CStr;
	CStr.Format( _T("%I64u"), Num ); //%i int //%u uint //%l long //%lu ulong //%ll longlong // %I64u Ulonglong //%5d int-right5
	return CStr;
}
#if 0
#include <inttypes.h>
CString CMFC_OPC_ExplorerDlg::LUintToCStr(uint64_t Num)
{
	CString CStr;
	CStr.Format( _T("%"PRId64""), Num ); //%i int //%u uint //%l long //%lu ulong //%ll longlong // %I64u Ulonglong //%5d int-right5
	return CStr;
}
CString CMFC_OPC_ExplorerDlg::LUintToCStr(uint32_t Num)
{
	CString CStr;
	CStr.Format( _T("%"PRId32""), Num ); //%i int //%u uint //%l long //%lu ulong //%ll longlong // %I64u Ulonglong //%5d int-right5
	return CStr;
}
#endif

CString CMFC_OPC_ExplorerDlg::FloatToCStr(float x, int count) // x - число, count - число знаков после запятой
{
	CString s_Result;
	CString s_Format("%.");
	s_Format = s_Format + LUintToCStr(count);
	s_Format = s_Format + _T("f");

	s_Result.Format( s_Format, x );
	return s_Result;
}
CString CMFC_OPC_ExplorerDlg::DoubleToCStr(double x, int count) // x - число, count - число знаков после запятой
{
	CString s_Result;
	CString s_Format("%.");
	s_Format = s_Format + LUintToCStr(count);
	s_Format = s_Format + _T("lf");

	s_Result.Format( s_Format, x );
	return s_Result;
}
CString CMFC_OPC_ExplorerDlg::DoubleToCStr(long double x, int count) // x - число, count - число знаков после запятой
{
	CString s_Result;
	CString s_Format("%.");
	s_Format = s_Format + LUintToCStr(count);
	s_Format = s_Format + _T("Lf");

	s_Result.Format( s_Format, x );
	return s_Result;
}



CString CMFC_OPC_ExplorerDlg::FormatOleDateTimeToCStr(COleDateTime odt, LPCTSTR strFormatString)
{
	CString strDateTime("");
	CString sFormatString(strFormatString);

	if (odt.m_dt == 0.0)
	{
		//strDateTime = "12/30/1899";

		// assumtion - we are free to do anything we want to the date since m_dt 
		// is set to 0.0
		//
		// consideration(s) - the user could potentially embed the very numbers 
		// we're going to use in our fake date, so we should account for them, 
		// maybe via a table or some sort of map.  At the same time, we could 
		// just ignore the possiblity

		// we require leading zeroes, so we have to strip the # character from 
		// the appropriate formatting commands...
		sFormatString.Replace(CString("%#d"), CString("%d"));
		sFormatString.Replace(CString("%#H"), CString("%H"));
		sFormatString.Replace(CString("%#I"), CString("%I"));
		sFormatString.Replace(CString("%#j"), CString("%j"));
		sFormatString.Replace(CString("%#m"), CString("%m"));
		sFormatString.Replace(CString("%#M"), CString("%M"));
		sFormatString.Replace(CString("%#S"), CString("%S"));
		sFormatString.Replace(CString("%#U"), CString("%U"));
		sFormatString.Replace(CString("%#w"), CString("%w")); // I know, we're technically replacing this twice
		sFormatString.Replace(CString("%#W"), CString("%W"));
		sFormatString.Replace(CString("%#y"), CString("%y"));
		sFormatString.Replace(CString("%#Y"), CString("%Y"));

		// the following formatting commands could potentially screw up our 
		// carefully laid plans, so we'll eliminate them.
		sFormatString.Replace(CString("%j"), CString(""));   // day of year
		sFormatString.Replace(CString("%w"), CString(""));   // day of week (0=sunday)

		// these formatting commands need to be intercepted since they
		// have specific values associated with the date we're trying 
		// to set (12/30/1899 00:00:00).
		sFormatString.Replace(CString("%a"), CString("Sat"));      // day of week - abbreviated
		sFormatString.Replace(CString("%A"), CString("Saturday")); // day of week - full
		sFormatString.Replace(CString("%p"), CString("PM"));   // am or pm
		
		// you'll see below why I set some fairly bizarre values here.
//		odt.SetDateTime(1900, 12, 29, 23, 59, 59);
		COleDateTime DateValidTill = COleDateTime::GetCurrentTime();
		
//		odt = COleDateTime::GetCurrentTime();
		strDateTime = odt.Format(sFormatString);
		// this one's easy...
		if (sFormatString.Find(CString("%Y")) >= 0 && strDateTime.Find(CString("1900")) >= 0)
		{
			strDateTime.Replace(CString("1900"), CString("1899"));
		}
		// if the programmer wants a 2-digit year, make a reasonable 
		// attempt to cover the most typical formatting possibilities
		if (sFormatString.Find(CString("%y")) >= 0)
		{
			// my first thought was to make sure the substring I wanted to 
			// replace was actually in the string, but that would have been 
			// inefficient (two "find" operations, followed by the actual 
			// replacement).  So, I figured I'd just call replace() for all 
			// of the possibilities
			strDateTime.Replace(CString("-00"), CString("-99"));
			strDateTime.Replace(CString("/00"), CString("/99"));
			strDateTime.Replace(CString(",00"), CString(",99"));
			strDateTime.Replace(CString(".00"), CString(".99"));
			strDateTime.Replace(CString("00"), CString("99")); // always do this one last
		}
		// now replace the hour (23 = 11pm)
		strDateTime.Replace(CString("23"), CString("00"));
		// the minutes and seconds
		strDateTime.Replace(CString("59"), CString("00"));
		// and finally the day of the month - we do this last to account for 
		// non-standard formatting possibilities
		strDateTime.Replace(CString("29"), CString("30"));
	}
	else
	{
		strDateTime = odt.Format(strFormatString);
	}

	return strDateTime;
}



///////////////////////////////////////////////////////////////////////////////////////
// Методы обработки сообщений об ошибках //////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////

void CMFC_OPC_ExplorerDlg::LogSessionStatus(CString Step, CString ParamsOrErrReason, CString ErrCode, bool bError)
{
	CString LogFileName("MFC_OPC_Explorer_Log.txt");
	//CString LogFileName(sCurProtName);

	// !!!!! ПРОТОКОЛИРУЕМ ОШИБКУ И ЕЁ ПРИЧИНУ
	int ii=0;
	CFile lfile;
	CStringA lLine;
	//CString lLine;
	const int l_FunName_len=20;
	const int l_Params_len=100;//максимальная длина записи
	char lOrderNum[10];

	BOOL lOk = lfile.Open(LogFileName, CFile::modeCreate|CFile::modeNoTruncate|CFile::modeReadWrite);
	if(!lOk)
	{	//"Невозможно открыть файл протокола, создается резервный файл протокола";
		//OpenTmpFile(LogFileName,lfile);
	}

	if(lOk && ((HANDLE)lfile!=INVALID_HANDLE_VALUE))
	{
		lfile.SeekToEnd();

		//признак нумерации строк протокола
		if(lfile.GetLength()>5)
		{
			ii = lfile.Seek(-8,CFile::current);
			lfile.Read(lOrderNum,11);
			ii = atoi(lOrderNum);
			if(ii>999998)
				strcpy(lOrderNum,"000000");
			else
				sprintf(lOrderNum,"%06i",ii+1);
		}
		else 
		{sprintf(lOrderNum,"%06i",0);}

		//wchar_t* tmp=NULL;
		char* tmp=NULL;

		//////////////////////////////////////////////////////////////////////
		GetLocalTime(&theApp.m_CurTime);	//Инициализируем системное время
		//lLine.Format(_T("|SOU|%02i.%02i.%04i | %02i:%02i:%02i|"),
		lLine.Format(("|OPC|%04i.%02i.%02i | %02i:%02i:%02i:%03i|"),
			theApp.m_CurTime.wYear, theApp.m_CurTime.wMonth, theApp.m_CurTime.wDay,
			theApp.m_CurTime.wHour, theApp.m_CurTime.wMinute, theApp.m_CurTime.wSecond, theApp.m_CurTime.wMilliseconds);

		lfile.Write(lLine,lLine.GetLength());

	//void BGetData::WriteFun2ComServerProt(CString vFunName, CString vParams, HRESULT vHR, int vResult)
	//{
//#if 1
		//////////////////////////////////////////////////////////////////////////
		lLine = Step;
		tmp = lLine.GetBufferSetLength(l_FunName_len);
		if(Step.GetLength()<l_FunName_len){
			//остаток строки забить пробелами
			memset(tmp+Step.GetLength(),0x20/*space char*/,l_FunName_len-Step.GetLength());
		}
		tmp[l_FunName_len] = 0;
		lLine.ReleaseBuffer();
		lfile.Write(lLine,l_FunName_len);
		lfile.Write("|",1);

		//////////////////////////////////////////////////////////////////////////
		DeleteNewStrings(ParamsOrErrReason);
		//if(bError)
		//{
		//	//lLine = _T("  ") + ParamsOrErrReason;
		//	lLine = ParamsOrErrReason;
		//}
		//else
		//{lLine = ParamsOrErrReason;}

		lLine = ParamsOrErrReason;


		tmp = lLine.GetBufferSetLength(l_Params_len);
		if(ParamsOrErrReason.GetLength()<l_Params_len)
		{
			//остаток строки забить пробелами
			memset(tmp+ParamsOrErrReason.GetLength(),0x20,l_Params_len-ParamsOrErrReason.GetLength());
		}
		tmp[l_Params_len] = 0;
		lLine.ReleaseBuffer();
		lfile.Write(lLine,l_Params_len);
		lfile.Write("|",1);

		//////////////////////////////////////////////////////////////////////////
		//lLine.Format("%10i",vHR);
		if(bError)
		{
			lLine = ErrCode;
			if(lLine.GetLength()<10){
				//остаток строки забить пробелами
				memset(tmp+lLine.GetLength(),0x20/*space char*/,l_FunName_len-lLine.GetLength());
			}
			tmp[10] = 0;
		}
		else
		{lLine.Format("%10i",0);}

		lfile.Write(lLine,10);
		lfile.Write("|",1);

		//////////////////////////////////////////////////////////////////////////
		lfile.Write("\r\n",2);
		lfile.Close();
	//#else
//#endif
	//}
	}
}



///////////////////////////////////////////////////////////////////////////////////////
// Некоторые вспомогательные методы обработки строк ///////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////

void CMFC_OPC_ExplorerDlg::DeleteNewStrings(CString& CheckStr)
{
	if(CheckStr.GetLength()>0)
	{
		CString TwoWCh("");
		for(unsigned long long i=0; i<CheckStr.GetLength(); i++)
		{
			TwoWCh.Empty();
			TwoWCh += CheckStr[i];
			//TwoWCh += CheckStr[i+1];
			if(TwoWCh == _T("\n"))
			{
				CheckStr.Delete(i,1);//2);
				i--;//i--;
				CheckStr +=(_T(" ")); // чтобы 2 строчки "не срослись"
			}
			//if(i==(CheckStr.GetLength()-1))//-2))
			//{break;}
		}
	}
}

void CMFC_OPC_ExplorerDlg::DeleteNewStringsA(CStringA& CheckStr)
{
	if(CheckStr.GetLength()>0)
	{
		CStringA TwoWCh("");
		for(unsigned long long i=0; i<CheckStr.GetLength(); i++)
		{
			TwoWCh.Empty();
			TwoWCh += CheckStr[i];
			//TwoWCh += CheckStr[i+1];
			if(TwoWCh == "\n")
			{
				CheckStr.Delete(i,1);//2);
				i--;//i--;
				CheckStr +=(" "); // чтобы 2 строчки "не срослись"
			}
			//if(i==(CheckStr.GetLength()-1))//-2))
			//{break;}
		}
	}
}

void CMFC_OPC_ExplorerDlg::DeleteLastSymbols(CString& Str, long NumOfSymbols)
{
	for(long i=0; i<NumOfSymbols; i++)
	{
		if(Str.GetLength()>0)
		{
			long i = Str.GetLength() - 1;
			//if(  (Str[i] != ' ')  &&  ( Str[i] != CString("") ) )
			Str.Delete(i);
		}
		else
		{break;}
		//return Str;
	}
}







void CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadCache()
{
	// TODO: добавьте свой код обработчика уведомлений
	int CheckRadio = GetCheckedRadioButton(IDC_BTN_TAG_READ_CACHE, IDC_BTN_TAG_READ_DEVICE);
}


void CMFC_OPC_ExplorerDlg::OnBnClickedBtnTagReadDevice()
{
	int CheckRadio = GetCheckedRadioButton(IDC_BTN_TAG_READ_CACHE, IDC_BTN_TAG_READ_DEVICE);
	// TODO: добавьте свой код обработчика уведомлений
}
