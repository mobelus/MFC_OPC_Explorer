#pragma once

#include "StdAfx.h"

// ��������� ��� �������� �������� ��� �������� �� �����
// � ��������� � ��������� ����� � �� ������� ��� ��������� � ������� ������� ���������

struct TagProps
{
public:
	long TagAccessRights;	// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
	CString TagEUUnits;		// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
	CString TagDescription;	// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
	CString TagProductName;	// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
	CString TagQuantity;	// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
	CString TagSiteDescr;	// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
	unsigned long TagSiteID;// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
};


struct TagStruct
{
public:
	int		  TagGroup;		// ����� � OPC-������� // ��������	// ��� ��������� � ����� �� �� ��������
	CString	  TagName;		// ����� � OPC-������� // ��������� // ��� ����������� ������� (������)
	CString	  TagFullName;	// ����� � OPC-������� // ��������� // ��� ����������� ������� (������)
	VARTYPE	  TagVrntType;	// ����� � OPC-������� // ��������� // ��� ����������� ������� (������)
	VARIANT	  TagVrntValue; // ����� � OPC-������� // ��������� // ��������, ��� ����� ����������� ���������
	CString	  TagCStrValue;	// ����� � OPC-������� // ��������� // ��������, ��� ����� ����������� ���������
	WORD	  TagQuality;	// ����� � OPC-������� // ��������� // ��������, ��� ����� ����������� ���������
	FILETIME  TagTimeStamp;	// ����� � OPC-������� // ��������� // ��������, ��� ����� ����������� ���������
//	OPCHANDLE TagHandle;	// ???

	TagProps TagProperties;	//

	//*
	TagStruct& operator=(const TagStruct& s)
	{
		(*this).TagGroup     = s.TagGroup;
		(*this).TagName      = s.TagName;
		(*this).TagFullName  = s.TagFullName;
		(*this).TagVrntType  = s.TagVrntType;
		(*this).TagVrntValue = s.TagVrntValue;
		(*this).TagCStrValue = s.TagCStrValue;
		(*this).TagQuality   = s.TagQuality;
		(*this).TagTimeStamp = s.TagTimeStamp;
		//(*this).TagHandle	 = s.TagHandle;
		(*this).TagProperties= s.TagProperties;
		return (*this);
	}
	//*/

};

/*
template <class Type>
void SetCStrStructElem(Type a) {
	a < b ? a : b;
}
*/

