#include "stdafx.h"
#include "UDWorkbook.h"
#include "UDTool.h"

namespace UDExcel{

CUDSheets CUDWorkbook::GetSheets()
{ 
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
 
	InvokeHelper(OLESTR("Sheets"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return  vResult.pdispVal;
}

void CUDWorkbook::SaveAs( string sPath )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_BSTR;//
	x.bstrVal  = ::SysAllocString(a2w(sPath.c_str()))    ;//   
	DISPID dispidNamed = DISPATCH_METHOD;
	// 		
	dpNoArgs.cArgs      = 1; 		
//	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
//	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	 
	
	InvokeHelper(OLESTR("SaveAs"), DISPATCH_METHOD, dpNoArgs, vResult);
}

void CUDWorkbook::Save()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
 	
	InvokeHelper(OLESTR("Save"), DISPATCH_METHOD, dpNoArgs, vResult);
}

void CUDWorkbook::Close()
{	
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;	
	InvokeHelper(OLESTR("Close"), DISPATCH_METHOD, dpNoArgs, vResult);
}
}