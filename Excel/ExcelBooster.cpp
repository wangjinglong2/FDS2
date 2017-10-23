#include "StdAfx.h"
#include "ExcelBooster.h"


MSExcelBooster::MSExcelBooster()
{
	CoInitialize(NULL);
}

MSExcelBooster::~MSExcelBooster()
{
	m_Excel.ReleaseDispatch();
	CoUninitialize();
}

BOOL MSExcelBooster::InitExcelCOM( bool bNewOneWorkSheets /*= true*/ )
{
	CLSID clsid;
	CLSIDFromProgID(L"Excel.Application", &clsid);  
	IUnknown *pUnk = NULL;
	IDispatch *pDisp = NULL;

	HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pDisp);
	if(FAILED(hr))
	{ 		
		MessageBox(0, _T("请检查是否已经安装EXCEL！"), _T(""), MB_OK);
		return 0;
	}
	//HRESULT hr = GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);
	//if(SUCCEEDED(hr)) 
	//{
	//	hr = pUnk->QueryInterface(IID_IDispatch, (void **)&pDisp);
	//}
	//if (!pDisp)
	//{
	//	return FALSE;
	//}
	//if (pUnk) 
	//	pUnk->Release();

	m_Excel.AttachDispatch(pDisp);

	//创建Excel 2000服务器(启动Excel)
	//if(!m_Excel.CreateDispatch(_T("Excel.Application"))) 
	//{
	//	AfxMessageBox(_T("无法启动Excel服务器!"));
	//	return FALSE;
	//}
	m_workbooks.AttachDispatch(m_Excel.GetWorkbooks());
	return TRUE;
}

void MSExcelBooster::ReleaseExcelCom()
{
	m_Excel.ReleaseDispatch();
	CoUninitialize();
}

void MSExcelBooster::OpenOneWorkSheets( CString sFileName )
{
	try
	{
		LPDISPATCH lpDisp;
		const COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR); 
		lpDisp = m_workbooks.Open(sFileName,      
			covOptional, covOptional, covOptional, covOptional, covOptional,
			covOptional, covOptional, covOptional, covOptional, covOptional,
			covOptional, covOptional );
		m_workbook.AttachDispatch(lpDisp);
	}
	catch (COleDispatchException *e)
	{
		e->ReportError();
		return;
	}
}

BOOL MSExcelBooster::SetCurWorkSheet( int nCurSheet )
{
	m_worksheets.AttachDispatch(m_workbook.GetWorksheets());
	try
	{
		COleVariant ole((short)nCurSheet);
		m_worksheet.AttachDispatch(m_worksheets.GetItem(ole));
	}
	catch (COleDispatchException *e)
	{
		e->ReportError();
		return FALSE;
	}
	m_worksheet.Activate();
	return TRUE;
}

void MSExcelBooster::MoveTo( const long nRowIndex,const long nColunIndex )
{
	if (nRowIndex<=0 || nColunIndex<=0)
	{
		return;
	}
	int nShi = (nColunIndex-1)/26;
	int nGe = (nColunIndex-1)%26;
	CString strShi;
	CString szCurCell;
	if (nShi>0)
	{
		strShi.Format(_T("%c"),'A'+nShi-1);
	}
	szCurCell.Format(_T("%c%d"),'A'+nGe,nRowIndex);
	szCurCell = strShi+szCurCell;

	COleVariant vRow(szCurCell);
	MSExcel::Range	range;
	LPDISPATCH disp = m_worksheet.GetRange(vRow,vRow);
	if (!disp)
	{
		return;
	}
	range.AttachDispatch(disp);
	range.Activate();
	m_nCurrentColumnIndex = nColunIndex;
	m_nCurrentRowIndex = nRowIndex;
}

void MSExcelBooster::NextColumn()
{
	MoveTo(m_nCurrentRowIndex,m_nCurrentColumnIndex+1);
}

MSExcelBooster & MSExcelBooster::operator<<( const int value )
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	COleVariant newValue((long)value);
	range.SetItem(oleRow,oleColumn,newValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator<<( const short value )
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	COleVariant newValue(value);
	range.SetItem(oleRow,oleColumn,newValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator<<( const double value )
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	COleVariant newValue(value);
	range.SetItem(oleRow,oleColumn,newValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator<<( const TCHAR *sValue)
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	COleVariant newValue(sValue);
	range.SetItem(oleRow,oleColumn,newValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(int &value)
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	VARIANT	newvalue = range.GetItem(oleRow,oleColumn);
	value = newvalue.intVal;
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(long &value)
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	VARIANT	newvalue = range.GetItem(oleRow,oleColumn);
	value = newvalue.lVal;
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(short &value)
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	VARIANT	newvalue = range.GetItem(oleRow,oleColumn);
	value = newvalue.iVal;
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(double &value)
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	VARIANT	newvalue = range.GetItem(oleRow,oleColumn);
	value = newvalue.dblVal;
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(TCHAR *sValue)
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	VARIANT	newvalue = range.GetItem(oleRow,oleColumn);
	sValue = newvalue.bstrVal;
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(CString &sValue)
{
	LPDISPATCH disp = m_worksheet.GetUsedRange();
	MSExcel::Range	range;
	if (!disp)
	{
		return *this;
	}
	range.AttachDispatch(disp);
	range.Activate();
	long nRow = range.GetRow();
	long nColumn = range.GetColumn();
	COleVariant oleRow(nRow);
	COleVariant oleColumn(nColumn);
	VARIANT	newvalue = range.GetItem(oleRow,oleColumn);
	sValue = newvalue.bstrVal;
	NextColumn();
	return *this;
}

MSExcelBooster& MSExcelBooster::endl( MSExcelBooster &excel )
{
	
	return *this;
}