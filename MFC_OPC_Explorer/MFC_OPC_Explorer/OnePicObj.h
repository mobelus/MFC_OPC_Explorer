#pragma once

#include "StdAfx.h"

// Структуры для хранения объектов для орисовки на форме
// и Структуры с описанием Тегов и их свойств что привязаны к данному объекту отрисовки

struct TagProps
{
public:
	long TagAccessRights;	// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
	CString TagEUUnits;		// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
	CString TagDescription;	// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
	CString TagProductName;	// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
	CString TagQuantity;	// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
	CString TagSiteDescr;	// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
	unsigned long TagSiteID;// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
};


struct TagStruct
{
public:
	int		  TagGroup;		// нужен в OPC-клиенте // единожды	// раз запросили и более он не меняется
	CString	  TagName;		// нужен в OPC-клиенте // постоянно // для дальнейшего запроса (чтения)
	CString	  TagFullName;	// нужен в OPC-клиенте // постоянно // для дальнейшего запроса (чтения)
	VARTYPE	  TagVrntType;	// нужен в OPC-клиенте // постоянно // для дальнейшего запроса (чтения)
	VARIANT	  TagVrntValue; // нужен в OPC-клиенте // постоянно // параметр, что будет обновляться регулярно
	CString	  TagCStrValue;	// нужен в OPC-клиенте // постоянно // параметр, что будет обновляться регулярно
	WORD	  TagQuality;	// нужен в OPC-клиенте // постоянно // параметр, что будет обновляться регулярно
	FILETIME  TagTimeStamp;	// нужен в OPC-клиенте // постоянно // параметр, что будет обновляться регулярно
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

