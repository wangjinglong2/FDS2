#include "stdafx.h"
#include "UDSheets.h"
#include "UDTool.h"

namespace UDExcel{
 
CUDSheet CUDSheets::AddSheet()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 

	InvokeHelper(OLESTR("Add"), DISPATCH_METHOD, dpNoArgs, vResult);
	
	return vResult.pdispVal;
}

CUDSheet CUDSheets::GetSheet( string sName )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_BSTR; 
	x.bstrVal  = ::SysAllocString(a2w(sName.c_str()))    ;
	DISPID dispidNamed = DISPATCH_PROPERTYGET;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
 
	InvokeHelper(OLESTR("Sheets"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return vResult.pdispVal;
}

CUDSheet CUDSheets::GetItem( int nIndex )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_I4;  
	x.intVal  = nIndex   ;
	DISPID dispidNamed = DISPATCH_PROPERTYGET;
	// 		
	dpNoArgs.cArgs      = 1; 		
//	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
//	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
		
	InvokeHelper(OLESTR("Item"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return vResult.pdispVal;
}

int CUDSheets::GetCount()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Count"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return vResult.intVal;
}


}