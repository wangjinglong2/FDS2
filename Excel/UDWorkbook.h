#pragma once

#include "UDSheets.h"
#include "UDIDispath.h"

namespace UDExcel{
class CUDWorkbook:public CUDIDispath
{
public:
 
	CUDWorkbook(CUDWorkbook& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; }
	CUDWorkbook(IDispatch*pDocs){ this->m_lpDispatch = pDocs; }
	CUDWorkbook& operator = (CUDWorkbook& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDWorkbook(){ /*if(!IsNull())m_pDocs->Release();*/ }

	CUDSheets GetSheets();
	void      SaveAs(string sPath);
	void      Save();
	void      Close();
};
}