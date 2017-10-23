#pragma once

#include "UDWorkbook.h"
#include "UDIDispath.h"
namespace UDExcel{
class CUDWorkbooks:public CUDIDispath
{
public:
	 
	CUDWorkbooks(CUDWorkbooks& Workbooks){ this->m_lpDispatch = Workbooks.m_lpDispatch; }
	CUDWorkbooks(IDispatch*pWorkbooks){ this->m_lpDispatch = pWorkbooks; }
	CUDWorkbooks& operator = (CUDWorkbooks& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDWorkbooks(){ /*if(!IsNull())m_pDocs->Release();*/ }


	BOOL        IsNull(){ return m_lpDispatch==0; }
	CUDWorkbook AddWorkbook(); 
	CUDWorkbook Open(string sPath);
};
}
