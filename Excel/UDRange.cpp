#include "stdafx.h"
#include "UDRange.h"
#include "UDTool.h"
namespace UDExcel{
  void CUDRange::SetValue( string sText )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt   = VT_BSTR;//
	  x.bstrVal  = ::SysAllocString(a2w(sText.c_str()))    ;//   
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;

	  InvokeHelper(OLESTR("Value"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }
  
  void CUDRange::Merge()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult; 

	  InvokeHelper(OLESTR("Merge"), DISPATCH_METHOD, dpNoArgs, vResult);
  }
  
  void CUDRange::UnMerge()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	   
	  InvokeHelper(OLESTR("UnMerge"), DISPATCH_METHOD, dpNoArgs, vResult);
  }
  
  void CUDRange::SetBKColor( int nColor )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	
	  OLECHAR FAR* szFunction = OLESTR("Interior");
	  DISPID  dispid_visiable;
	  HRESULT hr = m_lpDispatch->GetIDsOfNames (IID_NULL, &szFunction, 1, LOCALE_USER_DEFAULT, &dispid_visiable); 
	  hr = m_lpDispatch->Invoke(dispid_visiable, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &vResult, NULL, NULL);

	  IDispatch* iInterior = vResult.pdispVal;
	  {
		  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
		  VARIANT vResult;
		  
		  VARIANT x;
		  x.vt       = VT_I4;//
		  x.intVal   = nColor    ;//   
		  DISPID dispidNamed = DISPID_PROPERTYPUT;
		  // 		
		  dpNoArgs.cArgs      = 1; 		
		  dpNoArgs.cNamedArgs = 1;
		  dpNoArgs.rgvarg     = &x;
		  dpNoArgs.rgdispidNamedArgs = &dispidNamed;
		  
		  OLECHAR FAR* szFunction = OLESTR("Color");
		  DISPID  dispid_visiable;
		  HRESULT hr = iInterior->GetIDsOfNames (IID_NULL, &szFunction, 1, LOCALE_USER_DEFAULT, &dispid_visiable); 
		  hr = iInterior->Invoke(dispid_visiable, IID_NULL, LOCALE_USER_DEFAULT, DISPID_PROPERTYPUT, &dpNoArgs, &vResult, NULL, NULL);
	  }
  }

  CUDFont CUDRange::GetFont()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
 
	  InvokeHelper(OLESTR("Font"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	  return vResult.pdispVal;
  }
  
  void CUDRange::SetTextHAlign( int nAlign/*=0*/ )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt       = VT_I4;//
	  x.intVal   = nAlign;//::SysAllocString(a2w(sH_Text.c_str()))    ;//  
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;//DISPATCH_PROPERTYPUT
	  
	  InvokeHelper(OLESTR("HorizontalAlignment"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }


  void CUDRange::SetTextVAlign( int nAlign/*=TEXT_H_ALIGN_LEFT*/ )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt       = VT_I4;//
	  x.intVal   = nAlign;//::SysAllocString(a2w(sH_Text.c_str()))    ;//  
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;//DISPATCH_PROPERTYPUT
	  
	  InvokeHelper(OLESTR("VerticalAlignment"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }
  
  
  std::string CUDRange::GetValue()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;	  	 
	  
	  InvokeHelper(OLESTR("Value"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	  string sTxt;
	  switch(vResult.vt)
	  {
	  case VT_NULL:break;
	  case VT_BSTR: { sTxt = w2a(vResult.bstrVal); break; }
	  case VT_I4:{ sTxt = CI2A(vResult.intVal); break; }
	  case VT_R8:{ sTxt = CF2A(vResult.dblVal); break;}
	  }

	  return sTxt;
  }
  
  void CUDRange::SetSelect()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	 	  
	  InvokeHelper(OLESTR("Select"), DISPATCH_METHOD, dpNoArgs, vResult);
  }

  
  void CUDRange::ClearAll()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  InvokeHelper(OLESTR("Clear"), DISPATCH_METHOD, dpNoArgs, vResult);
  }
  
  
  double CUDRange::Top()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;	  	 
	  
	  InvokeHelper(OLESTR("Top"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	  
	  return vResult.dblVal;
  }
  
  double CUDRange::Left()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;	  	 
	  
	  InvokeHelper(OLESTR("Left"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
		  
	  return vResult.dblVal;
  }
  
  double CUDRange::Height()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;	  	 
	  
	  InvokeHelper(OLESTR("Height"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	  
	  return vResult.dblVal;
  }
  
  double CUDRange::Width()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;	  	 
	  
	  InvokeHelper(OLESTR("Width"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);
	  
	  return vResult.dblVal;
  }
  
  void CUDRange::ClearContents()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  InvokeHelper(OLESTR("ClearContents"), DISPATCH_METHOD, dpNoArgs, vResult);
  }
  
  void CUDRange::ClearFormats()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  InvokeHelper(OLESTR("ClearFormats"), DISPATCH_METHOD, dpNoArgs, vResult);
  }
  
  void CUDRange::ClearComments()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  InvokeHelper(OLESTR("ClearComments"), DISPATCH_METHOD, dpNoArgs, vResult);
  }
  
  int CUDRange::GetEndRow()
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;

	  VARIANT x; x.vt=VT_I4; x.intVal=-4162;
	  dpNoArgs.cArgs=1;
	  dpNoArgs.rgvarg = &x;
	  
	  InvokeHelper(OLESTR("End"), DISPATCH_PROPERTYGET, dpNoArgs, vResult);

	  DISPPARAMS dpNoArgs_ = {NULL, NULL, 0, 0};
	  CUDIDispath End = vResult.pdispVal;
	  End.InvokeHelper(OLESTR("Row"), DISPATCH_PROPERTYGET, dpNoArgs_, vResult);
	  
	  return vResult.intVal;
  }
//////////////////////////////////////////////////////////////////////////
  void CUDFont::SetColor( int nColor )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt       = VT_I4;//
	  x.intVal   = nColor    ;//   
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	   
	  InvokeHelper(OLESTR("Color"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }	
  
  void CUDFont::SetSize( int nSize )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt       = VT_I4;//
	  x.intVal   = nSize    ;//   
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;	   

	  InvokeHelper(OLESTR("Size"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }	
  
  void CUDFont::SetName( string sName )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt       = VT_BSTR;//
	  x.bstrVal = ::SysAllocString(a2w(sName.c_str()))    ;//  
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	  
	  InvokeHelper(OLESTR("Name"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }	
  
  void CUDFont::SetItalic( BOOL bItalic/*=TRUE*/ )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt       = VT_BOOL;//
	  x.boolVal  = bItalic;//  
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	  
	  InvokeHelper(OLESTR("Italic"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }
  
  void CUDFont::SetBold( BOOL bBold/*=TRUE*/ )
  {
	  DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	  VARIANT vResult;
	  
	  VARIANT x;
	  x.vt       = VT_BOOL;//
	  x.boolVal  = bBold;//  
	  DISPID dispidNamed = DISPID_PROPERTYPUT;
	  // 		
	  dpNoArgs.cArgs      = 1; 		
	  dpNoArgs.cNamedArgs = 1;
	  dpNoArgs.rgvarg     = &x;
	  dpNoArgs.rgdispidNamedArgs = &dispidNamed;
	  
	  InvokeHelper(OLESTR("Bold"), DISPID_PROPERTYPUT, dpNoArgs, vResult);
  }
  }