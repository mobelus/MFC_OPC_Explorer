
// MFC_OPC_Explorer.h : главный файл заголовка для приложения PROJECT_NAME
//

#pragma once

#ifndef __AFXWIN_H__
	#error "включить stdafx.h до включения этого файла в PCH"
#endif

#include "resource.h"		// основные символы


// CMFC_OPC_ExplorerApp:
// О реализации данного класса см. MFC_OPC_Explorer.cpp
//

class CMFC_OPC_ExplorerApp : public CWinApp
{
public:
	CMFC_OPC_ExplorerApp();

// Переопределение
public:
	virtual BOOL InitInstance();

//	Некоторые параметры использующиеся в разных частях приложения
	SYSTEMTIME m_CurTime;	//текущее системное время

	CString sIniFileName;
	CString sIniFilePath;

// Реализация

	DECLARE_MESSAGE_MAP()
};

extern CMFC_OPC_ExplorerApp theApp;