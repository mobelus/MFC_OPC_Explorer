
// MFC_OPC_Explorer.h : ������� ���� ��������� ��� ���������� PROJECT_NAME
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�������� stdafx.h �� ��������� ����� ����� � PCH"
#endif

#include "resource.h"		// �������� �������


// CMFC_OPC_ExplorerApp:
// � ���������� ������� ������ ��. MFC_OPC_Explorer.cpp
//

class CMFC_OPC_ExplorerApp : public CWinApp
{
public:
	CMFC_OPC_ExplorerApp();

// ���������������
public:
	virtual BOOL InitInstance();

//	��������� ��������� �������������� � ������ ������ ����������
	SYSTEMTIME m_CurTime;	//������� ��������� �����

	CString sIniFileName;
	CString sIniFilePath;

// ����������

	DECLARE_MESSAGE_MAP()
};

extern CMFC_OPC_ExplorerApp theApp;