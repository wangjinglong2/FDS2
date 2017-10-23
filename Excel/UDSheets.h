#pragma once

#include "UDSheet.h"
#include "UDIDispath.h"

namespace UDExcel{
class CUDSheets:public CUDIDispath
{
public: 
	CUDSheets(CUDSheets& Sheets){ this->m_lpDispatch = Sheets.m_lpDispatch; }
	CUDSheets(IDispatch*pSheets){ this->m_lpDispatch = pSheets; }
	CUDSheets& operator = (CUDSheets& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDSheets(){ /*if(!IsNull())m_pDocs->Release();*/ }
	
	
	BOOL        IsNull(){ return m_lpDispatch==0; }
	CUDSheet    GetActiveSheet();
	CUDSheet    AddSheet();
	CUDSheet    GetSheet(string sName);
	int         GetCount();
	CUDSheet    GetItem(int nIndex);//»ùÓÚ1	 
};
}