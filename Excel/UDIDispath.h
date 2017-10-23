#pragma once
#include <ole2.h> 

namespace UDExcel{  
class CUDIDispath
{ 
public:
	CUDIDispath();
	CUDIDispath(IDispatch*lpDispatch){ this->m_lpDispatch = lpDispatch; }

	void         InvokeHelper(OLECHAR FAR*FunName, WORD wFlags, DISPPARAMS&dpNoArgs, VARIANT&vResult);
	BOOL         IsValid();

	IDispatch*   GetExcelApp();
protected:
	IDispatch*   m_lpDispatch;

};  

}
