#include "stdafx.h"
#include "UDSheet.h"
#include "UDTool.h"
namespace UDExcel{
void CUDSheet::SetName( string sName )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_BSTR;//
	x.bstrVal  = ::SysAllocString(a2w(sName.c_str()))    ;//   
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;

	InvokeHelper(OLESTR("Name"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
}

std::string CUDSheet::GetName()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
 
	InvokeHelper(OLESTR("Name"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return w2a(vResult.bstrVal);
}

string GetCellName(int nRow, int nCol)
{
	char c = nCol+64;
	string sCol; sCol += c;

	string sName; sName=sCol; sName+=CI2A(nRow);
	sName+=":";	sName+=sCol; sName+=CI2A(nRow);

	return sName;
}

string GetColName(int nStartCol, int nEndCol)
{
	char c = nStartCol + 64;
	char cc= nEndCol + 64;
	string sCol; sCol += c;
	string ssCol; ssCol += cc;
	
	string sName; sName=sCol; 
	sName+=":";	sName+=ssCol; 
	
	return sName;
}

string GetColName(int nCol)
{
	char c = nCol + 64;
	string sCol; sCol += c;
		
	return sCol;
}

string GetRowName(int nStartRow, int nEndRow)
{
	string sCol; sCol += CI2A(nStartRow);
	string ssCol; ssCol += CI2A(nEndRow);
	
	string sName; sName=sCol; 
	sName+=":";	sName+=ssCol; 
	
	return sName;
}

string GetRangeName(int nStartRow, int nEndRow, int nStartCol, int nEndCol)
{
	char r1 = nStartCol + 64;
	char r2 = nEndCol + 64;

	string sR1; sR1 += r1;
	string sR2; sR2 += r2;

	sR1 += CI2A(nStartRow);
	sR2 += CI2A(nEndRow);

	return string(sR1+":"+sR2);
}

CUDRange CUDSheet::GetRange( int nRow, int nCol )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;

	VARIANT x;
	x.vt   = VT_BSTR;//

	string sName = GetCellName(nRow, nCol);
	x.bstrVal  = ::SysAllocString(a2w(sName.c_str()))    ;//   
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 0;
	dpNoArgs.rgvarg     = &x;
//	dpNoArgs.rgdispidNamedArgs = &dispidNamed;

 
	InvokeHelper(OLESTR("Range"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return  vResult.pdispVal;//获取到文档集
}
 
CUDRange CUDSheet::GetRange( int nStartRow, int nEndRow, int nStartCol, int nEndCol )
{
	string sRangeName = GetRangeName(nStartRow, nEndRow, nStartCol, nEndCol);
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	VARIANT x;
	x.vt   = VT_BSTR;//
	 
	x.bstrVal  = ::SysAllocString(a2w(sRangeName.c_str()))    ;//   
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 0;
	dpNoArgs.rgvarg     = &x;
	//	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	 

	InvokeHelper(OLESTR("Range"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return  vResult.pdispVal;//获取到文档集
}

void CUDSheet::Delete()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 

	InvokeHelper(OLESTR("Delete"), DISPATCH_METHOD, dpNoArgs, vResult);
}

void CUDSheet::SetVisale( BOOL bVisable )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	VARIANT x;
	x.vt   = VT_BOOL;//
	x.boolVal = bVisable; 
		
	DISPID dispidNamed = DISPID_PROPERTYPUT;

	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	dpNoArgs.cNamedArgs = 1;

	InvokeHelper(OLESTR("Visible"), DISPATCH_PROPERTYPUT, dpNoArgs, vResult);
}

void CUDSheet::SetActive()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 

	InvokeHelper(OLESTR("Select"), DISPATCH_METHOD, dpNoArgs, vResult);
}

void CUDSheet::CopyInsterBefore()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	
	VARIANT x;
	x.vt   = VT_DISPATCH;//
	x.pdispVal = m_lpDispatch; 
	
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.rgvarg     = &x; 
	InvokeHelper(OLESTR("Copy"), DISPATCH_METHOD, dpNoArgs, vResult);
}
 

CUDColumns CUDSheet::GetColumns( int nStartCol, int nEndCol )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;

	VARIANT x;
	x.vt   = VT_BSTR;//

	string sName = GetColName(nStartCol, nEndCol);
	x.bstrVal  = ::SysAllocString(a2w(sName.c_str()))    ;//   
 
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.rgvarg     = &x; 

	InvokeHelper(OLESTR("Select"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return vResult.pdispVal;
}

CUDRows CUDSheet::GetRows( int nStartRow, int nEndRow )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	VARIANT x;
	x.vt   = VT_BSTR;//
	
	string sName = GetRowName(nStartRow, nEndRow);
	x.bstrVal  = ::SysAllocString(a2w(sName.c_str()))    ;//   
	 
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 0;
	dpNoArgs.rgvarg     = &x;
 	
	InvokeHelper(OLESTR("Rows"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return vResult.pdispVal;
}

CUDPictures CUDSheet::GetPictures()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	 	
	InvokeHelper(OLESTR("Pictures"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	return vResult.pdispVal;
}

UDExcel::CUDShapes CUDSheet::GetShapes()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Shapes"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	
	return vResult.pdispVal;
}

void CUDSheet::MoveBefore( CUDSheet&sheet )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	
	VARIANT x;
	x.vt   = VT_DISPATCH;//
	x.pdispVal = sheet.m_lpDispatch; 
	
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.rgvarg     = &x; 
	InvokeHelper(OLESTR("Move"), DISPATCH_METHOD, dpNoArgs, vResult);
}

int CUDSheet::GetEndRow( int nCol )
{
	CUDRange rg = GetRange(65536, 1);
	return rg.GetEndRow();
}

void CUDColumns::SetWidth( double dWidth )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;

	VARIANT x;
	x.vt   = VT_I4;//
	x.lVal = dWidth;	
 
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;	 

	InvokeHelper(OLESTR("ColumnWidth"), DISPATCH_PROPERTYPUT, dpNoArgs, vResult);
}

void CUDColumns::Delete()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;

	InvokeHelper(OLESTR("Delete"), DISPATCH_METHOD, dpNoArgs, vResult);
}

void CUDColumns::SetVisale( BOOL bVisable/*=true*/ )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	VARIANT x;
	x.vt   = VT_BOOL;//
	x.boolVal = bVisable; 
	
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	dpNoArgs.cNamedArgs = 1;
	
	InvokeHelper(OLESTR("Visible"), DISPATCH_PROPERTYPUT, dpNoArgs, vResult);
}
 
void CUDColumns::SetSelect()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Select"), DISPATCH_METHOD, dpNoArgs, vResult);
}

void CUDRows::SetHeight( double dWidth )
{	
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT x;
	x.vt   = VT_I4;//
	x.lVal = dWidth;	
	
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	 
	VARIANT vResult;
	InvokeHelper(OLESTR("RowHeight"), DISPATCH_PROPERTYPUT, dpNoArgs, vResult);
}

void CUDRows::Delete()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 

	InvokeHelper(OLESTR("Delete"), DISPATCH_METHOD, dpNoArgs, vResult);
}

void CUDRows::SetVisale( BOOL bVisable/*=true*/ )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	VARIANT x;
	x.vt   = VT_BOOL;//
	x.boolVal = bVisable; 
	
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	dpNoArgs.cNamedArgs = 1;
	
	InvokeHelper(OLESTR("Visible"), DISPATCH_PROPERTYPUT, dpNoArgs, vResult);
}

void CUDRows::SetSelect()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Select"), DISPATCH_METHOD, dpNoArgs, vResult);
}

CUDPicture CUDPictures::InsertPicture(string sPath)
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	VARIANT x;
	x.vt   = VT_BSTR;//
	x.bstrVal = ::SysAllocString(a2w(sPath.c_str()));	
	
	// 		
	dpNoArgs.cArgs      = 1; 
	dpNoArgs.rgvarg     = &x;
	 

	InvokeHelper(OLESTR("Insert"), DISPATCH_METHOD, dpNoArgs, vResult);

	return vResult.pdispVal;
}

int CUDPictures::GetCount()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Count"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	
	return vResult.intVal;
}

CUDPicture CUDPictures::GetItem( int nIndex )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_I4;  
	x.intVal  = nIndex   ;
	DISPID dispidNamed = DISPATCH_PROPERTYGET;
	// 		
	dpNoArgs.cArgs      = 1; 	
	dpNoArgs.rgvarg     = &x;
	
	InvokeHelper(OLESTR("Item"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	
	return vResult.pdispVal;
}

void CUDPicture::SetWidth( double dWidth )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT x;
	x.vt   = VT_I4;//
	x.lVal = dWidth;	
	
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
	VARIANT vResult;
	InvokeHelper(OLESTR("Width"), DISPATCH_PROPERTYPUT, dpNoArgs, vResult);
}

void CUDPicture::SetHeight( double dHeight )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT x;
	x.vt   = VT_I4;//
	x.lVal = dHeight;	
	
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
	VARIANT vResult;
	InvokeHelper(OLESTR("Height"), DISPATCH_PROPERTYPUT, dpNoArgs, vResult);
}

void CUDPicture::SetName(string sName)
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_BSTR;//
	x.bstrVal  = ::SysAllocString(a2w(sName.c_str()))    ;//   
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
	InvokeHelper(OLESTR("Name"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
}

std::string CUDPicture::GetName()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Name"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	
	return w2a(vResult.bstrVal);
}

void CUDPicture::Delete()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	
	InvokeHelper(OLESTR("Delete"), DISPATCH_METHOD, dpNoArgs, vResult);
}

CUDShapeRange CUDPicture::GetShapeRange()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;

	InvokeHelper(OLESTR("ShapeRange"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);	
	return vResult.pdispVal;
}

int CUDShapes::GetCount()
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	InvokeHelper(OLESTR("Count"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	return vResult.intVal;
}

CUDShapeRange CUDShapes::GetRange( int nIndex )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_I4;  
	x.intVal  = nIndex   ;
	DISPID dispidNamed = DISPATCH_PROPERTYGET;
	// 		
	dpNoArgs.cArgs      = 1; 	
	dpNoArgs.rgvarg     = &x;
//	dpNoArgs.cNamedArgs = 1;
//	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
	InvokeHelper(OLESTR("Range"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	
	return vResult.pdispVal;
}

UDExcel::CUDPicture CUDShapes::AddPicture( string sPath, int Left, int Top, int Width, int Height )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult; 
	VARIANT vParam[7];
	
	VARIANT x;
	x.vt   = VT_BSTR;//
	x.bstrVal = ::SysAllocString(a2w(sPath.c_str()));	
	 
	
	vParam[6] = x;
	vParam[5].vt = VT_I4;
	vParam[5].boolVal = -1;
	vParam[4].vt = VT_I4;
	vParam[4].boolVal = 1;

	vParam[3].vt = VT_I4;
	vParam[3].intVal = Left;
	vParam[2].vt = VT_I4;
	vParam[2].intVal = Top;
	vParam[1].vt = VT_I4;
	vParam[1].intVal = Width;
	vParam[0].vt = VT_I4;
	vParam[0].intVal = Height;
	
	// 		
	dpNoArgs.cArgs      = 7; 
	dpNoArgs.rgvarg     = vParam;
	
	
	
	InvokeHelper(OLESTR("AddPicture"), DISPATCH_METHOD, dpNoArgs, vResult);
	
	return vResult.pdispVal;
}
 
/*
设置标准横高和列宽。请先插入一张图片，根据实际需要设置。         
            Selection.RowHeight = 89.25                
			Selection.ColumnWidth = 18.88
‘选择并调整图片位置  
			ActiveSheet.Pictures.Insert(fn).Select                   
			Selection.ShapeRange.IncrementLeft 2                  
			Selection.ShapeRange.IncrementTop 2 ’
取消纵横比，以便设置所有图片大小一致。   
			Selection.ShapeRange.LockAspectRatio = 0                 
			Selection.ShapeRange.Width = 112                   
			Selection.ShapeRange.Height = 85
*/


void CUDShapeRange::SetLockAspectRatio( BOOL bLock/*=TRUE*/ )
{
	DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	
	VARIANT x;
	x.vt   = VT_I4;//参数类型 long
	x.lVal = bLock    ;//具体参数   
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	// 		
	dpNoArgs.cArgs      = 1; 		
	dpNoArgs.cNamedArgs = 1;
	dpNoArgs.rgvarg     = &x;
	dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	
	InvokeHelper(OLESTR("LockAspectRatio"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
}

 }