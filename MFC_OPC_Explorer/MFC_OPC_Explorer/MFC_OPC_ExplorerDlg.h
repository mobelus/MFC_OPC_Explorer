
// MFC_OPC_ExplorerDlg.h : файл заголовка
//

#pragma once
#include "afxwin.h"
#include "afxcmn.h"

#include "opcda.h"

#include "helpful.h"

#include <vector>
#include <list>
#include <map>

#include "OnePicObj.h"

//отображение "имя группы" --> "описатель группы" (Строка в Хэндл)
typedef std::map< std::wstring, OPCHANDLE > mapSH;
typedef std::list< std::wstring > listTags;

typedef std::map< DWORD/*PropID*/, CComVariant/*PropValue*/> mapTagProps;


/*
template <class T>
class second : public vector<T>
{
public:
	second() {  } 
	void SetValue(const T & val)
	{ value = val; }
};
*/

/*
class Dataset : public std::vector<CString>
{
	Dataset();
	~Dataset();

	public:

	// bruteforce way. // Dataset(vector<float> val){//for every val[i] call push_back(val[i]);}
};
*/

enum NumericindicarorLook {
	  TAG_GOOD = 0

	, TAG_ERROR
	, BAD_OTHER
	
	, BAD_ND
	, BAD_AV
	, BAD_KZ
};

enum OverlayPictureType {
	OVL_EMPTY = 0,

	OVL_NONSIGNALED,
	OVL_SIGNALED,
	OVL_UNDEF,
	OVL_BREAK,

	OVL_TIMEOUT = 10,

	OVL_NONSIGNALED_TO	= OVL_NONSIGNALED	+OVL_TIMEOUT,
	OVL_SIGNALED_TO		= OVL_SIGNALED		+OVL_TIMEOUT,
	OVL_UNDEF_TO		= OVL_UNDEF			+OVL_TIMEOUT,
	OVL_BREAK_TO		= OVL_BREAK			+OVL_TIMEOUT,
};


// диалоговое окно CMFC_OPC_ExplorerDlg
class CMFC_OPC_ExplorerDlg : public CDialogEx
{
// Создание
public:
	CMFC_OPC_ExplorerDlg(CWnd* pParent = NULL);	// стандартный конструктор

// Данные диалогового окна
	enum { IDD = IDD_MFC_OPC_EXPLORER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// поддержка DDX/DDV


////////////////////////////////////////////////////////////////
// 1 ///////////////////////////////////////////////////////////
//	Всё, что нам нужно для Подключения
	bool bOPCConnection;

	CString sOPCServerRmtSrvrNameOrIPAddr; // Remote Server Name
	CString sOPCServerProgID;
	CString sOPCServerConnTime;
	CString sOPCServerTagsTimeout;

////////////////////////////////////////////////////////////////
// 2 ///////////////////////////////////////////////////////////
    UINT nCheckReadWrite; // 0 = read // 1 = write

	bool bAddGroupTag;
	bool bReadWriteTag;
	int  dAddItemTag;
	bool bReadItemTag;
	bool bWriteItemTag;

	CString sGroupTagName;

	CString sTagName;
	CString sTagType;
	CString sTagValue;
	CString sTagQuality;

// 2.1 // Всё, что нужно для ОРС-Сервера

	CStdioFile err_wf, res_wf;

	CString FullErrMess;
	CString CurMethod;
	CString CurParams;
	CString CurErr;
	CString CurParamsAndErr;
	CString CurErrCode;
	//std::wstring ErrorString;
	CString ErrCStr;
	//void WriteErrLog(std::ofstream& err_file, CString ErrMess);
	CString GetFullErrMess(CString CurMethod, CString Params, CString ErrMess, CString ErrCode);
	CString lLine;


	void SetCmBxTagTypes();

////////////////////////////////////////////////////////////////
// 3 ///////////////////////////////////////////////////////////

	CImageList timl;

//////////////////////////////////////////////////////////////////////
// 1, 2, 3 ///////////////////////////////////////////////////////////
	
	void AddItemToListBox(CString sStrToAdd, int nNumOfLstBx);

///////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////	ОРС		///////////////////////////////////

// Свойства для работы с ОРС-сервером
	IOPCServer* pIOPCServer;	//pointer to IOPServer interface	//
	IOPCItemMgt *pIOPCItemMgt;	//pointer to IOPCItemMgt interface	//##ModelId=411B5A1A019C

	OPCHANDLE hServerGroup; // server handle to the group
	OPCHANDLE hServerItem;  // server handle to the item

	//OPCHANDLE opchMyItem;		//##ModelId=411B5A1A01AC
	DWORD dwRevisedUpdateRate;	//##ModelId=411B5A1A016E
	HRESULT hr;					//##ModelId=411B5A1A017D
	DWORD dwrUp;				//##ModelId=411B5A1A018C
	LONG lTimeBias;				//##ModelId=411B5A1A018E

	//bool bConstrOk;

	enum eError
	{
		E_ItemsPresent=1,	//Данный пункт присутствует в списке группы		//##ModelId=411B5A190391
		E_Items_NO_Present, //Данный пункт не присутствует в списке группы	//##ModelId=411B5A190392
		E_Group_Present,	//Группа с указанным именем существует			//##ModelId=411B5A190393
		E_Group_NO_Present, //Группа с указанным именем не существует		//##ModelId=411B5A19039F
		E_Group_No_Adrres,  //Отсутсвует адрес группы						//##ModelId=411B5A1903A0
		E_NO_Memory, //Не удалось выделить требуемый размер памяти память	//##ModelId=411B5A1903A1
		E_NO_AdrresIOPCAsyncIO2, //Отсутсвует адрес асинхронного интерфейса	//##ModelId=411B5A1903A2
		E_NO_AdrresIOPCSyncIO, //Отсутсвует адрес синхронного интерфейса	//##ModelId=411B5A1903AF
		E_NO_AdrresIOPCItemMgt,//Отсутсвует адрес  интерфейса значений		//##ModelId=411B5A1903B0
		E_NO_IOPCServer		//Не удается получить интерфес IOPCServer		//##ModelId=411B5A1903B1
	};


	time_t currentTime;	// текущее 
	int startTime;		// начало
	int timeElapsed;	// прошло


// Упакованные методы для работы с ОРС-сервером
	bool bNextStep;

	bool Connect(
				  LPCTSTR szHost
				, LPCTSTR szProgID
				, double dConnTime
//				, bool& bResult
				);
	bool DisConnect();


// Методы для работы с ОРС-сервером
//#define VCTR_INDX_ACTVSOU 0	// Активен/нет/СОУ
/*
	bool GetTagsValues(
						//std::vector<std::vector<CString>> v_Tags
						//, std::vector<CString>& v_Values
						std::vector<TagStruct>& v_Tags
					  );

	bool SetTagsValues( CString TagName, CString TagType, CString TagValue); //, CString UPorDN,	 CString PipeNUM, CString LorR,	int Command);
	//CString SetTagsValues( CString TagName, CString TagType, CString TagValue); //, CString UPorDN,	 CString PipeNUM, CString LorR,	int Command);
*/

	

#define VCTR_INDX_TAGNAME 0	// Имя тега
#define VCTR_INDX_TAGTYPE 1	// Тип тега
#define VCTR_INDX_INVRS	  2	// Тег инвертирован

	TagStruct CurTag;


	IOPCServer* InstantiateServer(CString OPCServProgID_CStr, CString RmtServNameOrIPAdr_CStr, bool& Result);
	int GetStatus(void);

	bool AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, OPCHANDLE& hServerGroup);
	bool WriteTag();
	bool ReadTag();

	bool AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, CString TagName, CString TagType);
	int AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, int iItemsNum, std::vector<std::vector<CString>> v_AllTags);
	int AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, TagStruct OneTag);

	bool ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue, WORD& wrdValue);
	bool ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, TagStruct& OneTag);
	bool WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT* varValue);
	bool WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, TagStruct OneTag);

// Методы обработки качества
	//int GetTypeOfQual(WORD wrdValue);
	//NumericindicarorLook GetTypeOfQual( TagStruct& val );
	NumericindicarorLook GetTypeOfQual( WORD wrdValue );
	OverlayPictureType GetSignalPicType( TagStruct& val );
	CString QualToBoolCStr(bool w, bool BigFirstLetter=true);
	CString QualToFullCStr(WORD wrdValue);

// Методы обработки Свойств
	CString OneTagPropsToFullCStr(IOPCServer* pIOPCServer, TagStruct OneTag);


	void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE hServerItem);
	void RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup);



	LPCWSTR lpcw;
	void TryBrowseBranch(IOPCBrowseServerAddressSpace *pBSAS, LPCWSTR lpprestr, LPCWSTR lpcw);
	void TryBrowseOneBranch(IOPCBrowseServerAddressSpace *pBSAS, LPCWSTR lpprestr, LPCWSTR lpcw);

	std::vector<std::vector<CString>> vv_OPCAddrSpace;
	std::map<int, CString> map_TagProps;
	std::map<int, CString> ::iterator TagPropsIter; // определяем итератор

	bool GetServerTags( /*[out]*/ std::list<std::wstring>& ref_list /*контейнер для заполнения*/ );
	bool GetSites( /*[out]*/ std::list<std::wstring>& ref_list /*контейнер для заполнения*/ );
	bool GetTagProperties(const std::wstring& strTag, mapTagProps& Props);
	bool GetTagProperty(const std::wstring& strTag, DWORD dwPropID, CComVariant& varProp);

	typedef std::list<std::wstring> ListTags;
/*
	typedef std::map<std::wstring, ItemValue, less_for_tags> MapTagValues;
	typedef std::map<DWORD, CComVariant> TagProps;
	typedef std::map<std::wstring,  TagProps, less_for_tags> MapTagProps;
*/

/*
  //По датчикам давления/расхода
	ListTags m_SensorTags;
	MapTagValues m_SensorValues;
	MapTagProps m_SensorTagProps;	// имя тэга --> список свойств

  //По датчикам открытия дверей
	ListTags 	 m_DoorTags;
	MapTagValues m_DoorStates;
	MapTagProps  m_DoorTagProps;		// имя тэга --> список свойств
*/


// Методы конвертации параметров
	VARTYPE CStrVarTypeToVariantType(CString VariantTYPE_CStr);
	_variant_t CStrToVariant(CString VariantVALUE_CStr, CString VariantTYPE_CStr);
	CString VariantTypeToCStrVarType(VARTYPE VariantTYPE);
	CString VariantValueToCStr(VARIANT varVal);

	bool CStrToBool(CString cstr);
	CString BoolToCStr(bool b, bool BigFirstLetter=true);
	CString LUintToCStr(unsigned int Num);
	CString LUintToCStr(long Num);
	CString LUintToCStr(unsigned long Num);
	CString LUintToCStr(__int32 Num); 	//CString LUintToCStr(int Num);
	CString LUintToCStr(__int64 Num); 	//CString LUintToCStr(long long Num);
	CString FloatToCStr(float x, int count=5); // x - число, count - число знаков после запятой
	CString DoubleToCStr(double x, int count=5); // x - число, count - число знаков после запятой
	CString DoubleToCStr(long double x, int count=5); // x - число, count - число знаков после запятой
	CString FormatOleDateTimeToCStr(COleDateTime odt, LPCTSTR strFormatString);

// Методы обработки сообщений об ошибках
	std::wstring GetErrorSTDWString(HRESULT hrErr);
	CString GetErrorCSTRString(HRESULT hrErr);
	void GetAndLogError(CString ErrMess);
	void LogSessionStatus(CString Step, CString ParamsOrErrReason, CString ErrCode, bool bError);

// Некоторые вспомогательные методы обработки строк
	void DeleteNewStrings(CString& CheckStr);
	void DeleteNewStringsA(CStringA& CheckStr);
	void DeleteLastSymbols(CString& Str, long NumOfSymbols);



////////////////////////////////////////////////////////////////
// Реализация
protected:
	HICON m_hIcon;

	// Созданные функции схемы сообщений
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	int a_m_z_;

	CEdit m_EditHost;
	CEdit m_EditProgID;
	CEdit m_EditTimeoutConnect;
	CEdit m_EditTimeoutTags;
	CButton m_BtnConnect;
	CStatic m_StaticConnectStatus;
	CButton m_BtnDisconnect;
	CListCtrl m_ListConnect;
//	CButton m_BtnConnectRefresh;
	CEdit m_EditGroupTagName;
	CButton m_BtnGroupTagName;
	CEdit m_EditTagName;
	CButton m_CheckTagType;
	CComboBox m_CmBxTagType;
	CEdit m_EditTagType;
	CEdit m_EditTagValue;
	CStatic m_StaticTagQual;
	CButton m_CheckTagReadWrite;
	CButton m_BtnTagReadWrite;
	int m_BtnTagReadFromCahche;
	int m_BtnTagReadFromDevice;
	CListBox m_LstBxTagRWProperties;
	CListCtrl m_ListTagReadWrite;
//	CButton m_LstBxTagReadWriteRefresh;
	CButton m_BtnGetSrvrAddrSpace;
	CTreeCtrl m_TreeGetSrvrAddrSpace;
	CListBox m_LstBxTagProperties;
	CListCtrl m_ListGetSrvrAddrSpace;
//	CButton m_LstBxGetSrvrAddrSpaceRefresh;

	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnEnChangeEditHost();
	afx_msg void OnEnChangeEditProgid();
	afx_msg void OnEnChangeEditTimeoutConnect();
	afx_msg void OnEnChangeEditTimeoutTags();
	afx_msg void OnBnClickedBtnConnect();
	afx_msg void OnBnClickedBtnDisconnect();
	afx_msg void OnBnClickedBtnConnectRefresh();
	afx_msg void OnEnChangeEditGroupTagName();
	afx_msg void OnBnClickedBtnGroupTagName();
	afx_msg void OnEnChangeEditTagName();
	afx_msg void OnBnClickedCheckTagType();
	afx_msg void OnCbnSelchangeCmbxTagType();
	afx_msg void OnEnChangeEditTagType();
	afx_msg void OnEnChangeEditTagValue();
	afx_msg void OnBnClickedCheckTagReadWrite();
	afx_msg void OnBnClickedBtnTagReadWrite();
	afx_msg void OnBnClickedBtnTagReadWriteRefresh();
	afx_msg void OnBnClickedBtnGetServerAddressSpace();
	afx_msg void OnTvnSelchangedTreeServerAddressSpace(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedBtnServerAddressSpaceRefresh();
	afx_msg void OnBnClickedBtnTagReadCache();
	afx_msg void OnBnClickedBtnTagReadDevice();
};
