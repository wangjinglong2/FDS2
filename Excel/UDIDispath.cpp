#include "StdAfx.h"
#include "UDIDispath.h"
#include <string>
#include "UDTool.h"
namespace UDExcel{
void CUDIDispath::InvokeHelper( OLECHAR FAR*FunName, WORD wFlags, DISPPARAMS&dpNoArgs, VARIANT&vResult )
{
	if(!IsValid())return;

	DISPID  dispID; 
	HRESULT hr = m_lpDispatch->GetIDsOfNames (IID_NULL, &FunName, 1, LOCALE_USER_DEFAULT, &dispID); 
	if(FAILED(hr)) 
	{
		char buf[512]="";

		std::string sFun = w2a(FunName);
        sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", sFun.c_str(), hr);
        
		throw CExcelExpection(buf);
    }

	hr = m_lpDispatch->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, wFlags, &dpNoArgs, &vResult, NULL, NULL);
	if(FAILED(hr)) 
	{
		char buf[512]=""; std::string sFun = w2a(FunName);
        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", sFun.c_str(), dispID, hr);
        throw CExcelExpection(buf);
    }
}

BOOL CUDIDispath::IsValid()
{
	return m_lpDispatch != 0;
}

IDispatch* CUDIDispath::GetExcelApp()
{ 
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 

	InvokeHelper(OLESTR("Application"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	
	return vResult.pdispVal;
}

CUDIDispath::CUDIDispath()
{
	m_lpDispatch=0;
}
}