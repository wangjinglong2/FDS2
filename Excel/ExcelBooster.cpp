#include "StdAfx.h"
#include "ExcelBooster.h"

COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  

MSExcelBooster::MSExcelBooster()
{
}

MSExcelBooster::~MSExcelBooster()
{
	m_workbooks.ReleaseDispatch();  
	m_workbook.ReleaseDispatch();  
	m_worksheets.ReleaseDispatch();  
	m_worksheet.ReleaseDispatch();  
	m_range.ReleaseDispatch();  
	m_font.ReleaseDispatch();  
	m_cell.ReleaseDispatch();  
	m_application.Quit();  
	m_application.ReleaseDispatch();  
	::CoUninitialize(); 
}

BOOL MSExcelBooster::InitExcelCOM()
{
	if (::CoInitialize( NULL ) == E_INVALIDARG)   
	{   
		AfxMessageBox(_T("初始化Com失败!"));   
		return FALSE;  
	}  
	if( !m_application.CreateDispatch(_T("Excel.Application")) )  
	{  
		AfxMessageBox(_T("无法创建Excel应用！"));  
		return FALSE;  
	} 
	return TRUE;
}

BOOL MSExcelBooster::OpenExcelBook( CString sFileName )
{
	CFileFind filefind;   
	if( !filefind.FindFile(sFileName) )   
	{   
		AfxMessageBox(_T("文件不存在"));  
		return TRUE;  
	}  
	LPDISPATCH lpDisp; //接口指针  
	m_workbooks=m_application.GetWorkbooks();  
	lpDisp = m_workbooks.Open(sFileName,      
		covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional );                 
	m_workbook.AttachDispatch(lpDisp);  
	m_worksheets=m_workbook.GetSheets();  
	m_worksheet=m_worksheets.GetItem(COleVariant((short)1));   
	return TRUE;  
}

BOOL MSExcelBooster::NewExcelBook()
{
	m_workbooks=m_application.GetWorkbooks();  
	m_workbook=m_workbooks.Add(covOptional);  
	m_worksheets=m_workbook.GetSheets();  
	m_worksheet=m_worksheets.GetItem(COleVariant((short)1));
	return TRUE;
}

BOOL MSExcelBooster::OpenExcelApp()
{
	m_application.SetVisible(TRUE);  
	m_application.SetUserControl(TRUE);  
	return TRUE;
}

BOOL MSExcelBooster::SaveExcel()  
{  
	m_workbook.SetSaved(TRUE);  
	return TRUE;
}

BOOL MSExcelBooster::SaveAsExcel(CString sFileName)  
{  
	m_workbook.SaveAs(COleVariant(sFileName),covOptional,covOptional,covOptional,
		covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional);
	return TRUE;
} 

BOOL MSExcelBooster::SetCellValue(int row, int col,CString sValue,int Align)  
{  
	m_range=m_worksheet.GetRange(COleVariant(GetCellPos(row,col)),COleVariant(GetCellPos(row,col)));  
	m_range.SetValue2(COleVariant(sValue));  
	m_cell.AttachDispatch((m_range.GetItem (COleVariant(long(1)), COleVariant(long(1)))).pdispVal);  
	m_cell.SetHorizontalAlignment(COleVariant((short)Align)); 
	return TRUE;
}  

CString MSExcelBooster::GetCellValue(int row, int col)  
{  
	m_range=m_worksheet.GetRange(COleVariant(GetCellPos(row,col)),COleVariant(GetCellPos(row,col)));  
	COleVariant rValue;  
	rValue=COleVariant(m_range.GetValue2());  
	rValue.ChangeType(VT_BSTR);  
	return CString(rValue.bstrVal);  
} 

BOOL MSExcelBooster::SetRowHeight(int row, CString height)  
{  
	int col = 1;  
	m_range=m_worksheet.GetRange(COleVariant(GetCellPos(row,col)),COleVariant(GetCellPos(row,col)));  
	m_range.SetRowHeight(COleVariant(height)); 
	return TRUE;
} 

CString MSExcelBooster::GetRowHeight(int row)  
{  
	int col = 1;  
	m_range=m_worksheet.GetRange(COleVariant(GetCellPos(row,col)),COleVariant(GetCellPos(row,col)));  
	VARIANT height = m_range.GetRowHeight();  
	CString strheight;  
	strheight = height;  
	return strheight;  
} 

BOOL MSExcelBooster::SetColumnWidth(int col,CString width)  
{  
	int row = 1;  
	m_range=m_worksheet.GetRange(COleVariant(GetCellPos(row,col)),COleVariant(GetCellPos(row,col)));  
	m_range.SetColumnWidth(COleVariant(width));  
	return TRUE;
} 

CString MSExcelBooster::GetColumnWidth(int col)  
{  
	int row = 1;  
	m_range=m_worksheet.GetRange(COleVariant(GetCellPos(row,col)),COleVariant(GetCellPos(row,col)));  
	VARIANT width = m_range.GetColumnWidth();  
	CString strwidth;  
	strwidth = width;  
	return strwidth;  
} 

CString MSExcelBooster::GetCellPos( int row, int col )   
{   
	CString strResult;  
	if( col > 26 )   
	{   
		strResult.Format(_T("%c%c%d"),'A' + (col-1)/26-1,'A' + (col-1)%26,row);  
	}   
	else   
	{   
		strResult.Format(_T("%c%d"), 'A' + (col-1)%26,row);  
	}   
	return strResult;  
} 

BOOL MSExcelBooster::SetCurWorkSheet( int nCurSheet )
{
	m_worksheet=m_worksheets.GetItem(COleVariant((short)nCurSheet));
	m_worksheet.Activate();
	return TRUE;
}

BOOL MSExcelBooster::MoveTo( const long nRowIndex,const long nColunIndex )
{
	if (nRowIndex<=0 || nColunIndex<=0)
	{
		return FALSE;
	}
	m_nCurrentColumnIndex = nColunIndex;
	m_nCurrentRowIndex = nRowIndex;
	return TRUE;
}

void MSExcelBooster::NextColumn()
{
	m_nCurrentColumnIndex ++;
}

MSExcelBooster & MSExcelBooster::operator<<( const int value )
{
	CString sTemp;
	sTemp.Format(_T("%d"),value);
	SetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex,sTemp);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator<<( const short value )
{
	CString sTemp;
	sTemp.Format(_T("%hd"),value);
	SetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex,sTemp);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator<<( const double value )
{
	CString sTemp;
	sTemp.Format(_T("%g"),value);
	SetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex,sTemp);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator<<( const TCHAR *sValue)
{
	CString sTemp;
	sTemp.Format(_T("%s"),sValue);
	SetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex,sTemp);
	NextColumn();
	return *this;
}

MSExcelBooster & MSExcelBooster::operator<<( const CString& sValue)
{
	SetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex,sValue);
	NextColumn();
	return *this;
}

MSExcelBooster & MSExcelBooster::operator>>(int &value)
{
	CString sValue = GetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex);
	value = _ttoi(sValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(long &value)
{
	CString sValue = GetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex);
	value = _ttol(sValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(short &value)
{
	CString sValue = GetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex);
	value = (short)_ttoi(sValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(double &value)
{
	CString sValue = GetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex);
	value = _tstof(sValue);
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(TCHAR *sValue)
{
	CString sTemp = GetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex);
	sValue = sTemp.GetBuffer(sTemp.GetLength());
	NextColumn();
	return *this;
}
MSExcelBooster & MSExcelBooster::operator>>(CString &sValue)
{
	sValue = GetCellValue(m_nCurrentRowIndex,m_nCurrentColumnIndex);
	NextColumn();
	return *this;
}

MSExcelBooster& endl( MSExcelBooster &excel )
{
	excel.SetCurrentRowIndex(excel.GetCurrentRowIndex()+1);
	excel.SetCurrentColumnIndex(1);
	return excel;
}

