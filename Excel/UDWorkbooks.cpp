#include "stdafx.h" 
#include "UDWorkbooks.h"
#include "UDTool.h"

namespace UDExcel{
CUDWorkbook CUDWorkbooks::AddWorkbook()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;	 

	InvokeHelper(OLESTR("Add"), DISPATCH_METHOD, dpNoArgs, vResult);
	
	return vResult.pdispVal;
}

CUDWorkbook CUDWorkbooks::Open( string sPath )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_BSTR; 
	x.bstrVal  = ::SysAllocString(a2w(sPath.c_str()))    ;
	DISPID dispidNamed = DISPATCH_METHOD;
	// 		
	dpNoArgs.cArgs      = 1; 		
//	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
//	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	 	
	InvokeHelper(OLESTR("Open"), DISPATCH_METHOD, dpNoArgs, vResult);

	return vResult.pdispVal;
}
}