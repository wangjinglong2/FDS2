// UDExcel.cpp: implementation of the CUDExcel class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "UDExcelApp.h"
#include "UDTool.h"

namespace UDExcel{
CUDExcelApp::CUDExcelApp()
{	
}

CUDExcelApp::~CUDExcelApp()
{
}

BOOL CUDExcelApp::InitServer()
{
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&m_lpDispatch);
	if(FAILED(hr))
	{ 		
		MessageBox(0, _T("请检查是否已经安装EXCEL！"), _T(""), MB_OK);
		return 0;
	}
	
	return 1;
}
BOOL CUDExcelApp::StopServer()
{
	if(!IsValid())return FALSE;
	
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Quit"), DISPATCH_METHOD, dpNoArgs, vResult);
	
	return TRUE;
}
 
CUDWorkbooks CUDExcelApp::GetWorkbooks()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	 
	InvokeHelper(OLESTR("Workbooks"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return  vResult.pdispVal;
}

CUDWorkbook CUDExcelApp::GetActiveWorkbook()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;	

	InvokeHelper(OLESTR("ActiveWorkbook"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return  vResult.pdispVal;//
}

void CUDExcelApp::SetVisable( BOOL bVisable )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_I4;//参数类型 long
	x.lVal = bVisable    ;//具体参数   
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
	InvokeHelper(OLESTR("Visible"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
}

CUDSheet CUDExcelApp::GetActiveSheet()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	 
	InvokeHelper(OLESTR("ActiveSheet"), DISPATCH_PROPERTYGET, dpNoArgs, vResult); 

	return  vResult.pdispVal;//
}

void CUDExcelApp::DisplayAlerts( BOOL bVisable )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_I4;//参数类型 long
	x.lVal = bVisable    ;//具体参数   
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
	InvokeHelper(OLESTR("DisplayAlerts"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
}

}