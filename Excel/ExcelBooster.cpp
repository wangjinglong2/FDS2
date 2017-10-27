#include "StdAfx.h"
#include "ExcelBooster.h"
#include "comutil.h"

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
	//m_application.Quit();  
	//m_application.ReleaseDispatch();  
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
	m_application.SetVisible(m_bIsVisible);
	m_application.SetDisplayAlerts(m_bDisplayAlerts);
	return TRUE;
}

BOOL MSExcelBooster::Quit()
{
	m_application.Quit();  
	m_application.ReleaseDispatch();
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
CString MSExcelBooster::GetCellPos(const CString& var)
{
	CString strResult; 
	m_range.AttachDispatch( m_worksheet.GetUsedRange() , TRUE);
	int		iRow,iColumn;
	MSExcel::Range	rowrange,colrange;
	rowrange = m_range.GetRows();
	colrange = m_range.GetColumns();
	iRow = rowrange.GetCount();
	iColumn = colrange.GetCount();
	rowrange.ReleaseDispatch();
	colrange.ReleaseDispatch();
	for (int i = 1; i <= iRow; i ++)
	{
		for (int j = 1; j <= iColumn; j ++)
		{
			CString sTemp;
			sTemp = GetCellValue(i,j);
			if (sTemp == var)
				return GetCellPos(i,j);
		}
	}
	return strResult;
}

BOOL MSExcelBooster::GetCellPos(const CString& var,int& iRow,int& iColumn)
{
	m_range.AttachDispatch( m_worksheet.GetUsedRange() , TRUE);
	int	iRowTemp,iColumnTemp;
	MSExcel::Range	rowrange,colrange;
	rowrange = m_range.GetRows();
	colrange = m_range.GetColumns();
	iRowTemp = rowrange.GetCount();
	iColumnTemp = colrange.GetCount();
	rowrange.ReleaseDispatch();
	colrange.ReleaseDispatch();
	for (int i = 1; i <= iRowTemp; i ++)
	{
		for (int j = 1; j <= iColumnTemp; j ++)
		{
			CString sTemp;			
			sTemp = GetCellValue(i,j);
			if (sTemp == var)
			{
				iRow = i;
				iColumn = j;
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL MSExcelBooster::MergeCell(const CString& sPlace1,const CString& sPlace2)
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)),TRUE);
	m_range.Merge(COleVariant((long)0));
	return TRUE;
}

BOOL MSExcelBooster::FillMerge(const CString &sPlace1,const CString &sPlace2,const CString &sText)
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)),TRUE);
	m_range.Merge(COleVariant((long)0));
	m_range.SetValue2(COleVariant(sText));
	return TRUE;
}

BOOL MSExcelBooster::InsertPicture(const CString& sPlace1,const CString& sPlace2,const CString& sPicPath,BOOL bZoom,const int nGap)
{
	MSExcel::Shapes    shapes = m_worksheet.GetShapes();
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)),TRUE);
	m_range.Merge(COleVariant(long(0)));
	float    fLeft, fTop, fWidth, fHeight;
	fLeft    = (float)m_range.GetLeft().dblVal;
	fTop    = (float)m_range.GetTop().dblVal;

	if ( bZoom )
	{
		// 图片进行缩放,保持与指定Range区域大小一致
		fWidth    = (float)m_range.GetWidth().dblVal;
		fHeight    = (float)m_range.GetHeight().dblVal;
	}
	else
	{
		// 图片保持原大小
		BITMAP bm;
		HBITMAP hbitmap;
		hbitmap = (HBITMAP) LoadImage(NULL, sPicPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
		GetObject(hbitmap, sizeof(bm), &bm);
		fWidth = (float)bm.bmWidth;
		fHeight = (float)bm.bmHeight;
	}
	shapes.AddPicture( sPicPath, false, true, fLeft, fTop, fWidth, fHeight );
	return TRUE;
}

BOOL MSExcelBooster::SetCurWorkSheet( int nCurSheet )
{
	m_worksheet=m_worksheets.GetItem(COleVariant((short)nCurSheet));
	m_worksheet.Activate();
	return TRUE;
}

BOOL MSExcelBooster::SetCurWorkSheet(int nCurSheet,const TCHAR *sSheetName)
{
	m_worksheet=m_worksheets.GetItem(COleVariant((short)nCurSheet));
	m_worksheet.SetName(sSheetName);
	m_worksheet.Activate();
	return TRUE;
}

BOOL MSExcelBooster::SetCurWorkSheet(const TCHAR *sSheetName)
{
	m_worksheet=m_worksheets.GetItem(COleVariant(sSheetName));
	m_worksheet.Activate();
	return TRUE;
}

BOOL MSExcelBooster::DeleteSheet(const short nItem)
{
	m_worksheet=m_worksheets.GetItem(COleVariant(nItem));
	m_worksheet.Delete();
	return TRUE;
}

BOOL MSExcelBooster::AddSheet()
{
	 m_worksheets.Add(vtMissing,vtMissing,_variant_t((long)1),vtMissing); 
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

BOOL MSExcelBooster::DeleteRow(int nRow)
{
	m_range.AttachDispatch( m_worksheet.GetRows() );
	m_range.AttachDispatch( m_range.GetItem( COleVariant( (short)nRow ), vtMissing ).pdispVal );
	m_range.Delete( COleVariant((long)ExcelBoosterBase::xlUp) );
	//m_nCurrentRowIndex--;
	return TRUE;

}
BOOL MSExcelBooster::DeleteRows( int nStartRow,int nEndRow )
{
	for (int i = nStartRow; i <= nEndRow; i ++)
		DeleteRow(nStartRow);
	return TRUE;
}

BOOL MSExcelBooster::InsertRow( int nStartRow,int nRowCnt )
{
	m_range.AttachDispatch( m_worksheet.GetRows() );
	m_range.AttachDispatch( m_range.GetItem( COleVariant( (short)nStartRow ), vtMissing ).pdispVal );
	for (int i = 1; i <= nRowCnt; i ++)
		m_range.Insert(COleVariant((long)ExcelBoosterBase::xlDown));
	return TRUE;
}

BOOL MSExcelBooster::CopyRow( int nFromRow,int nToRow )
{
	MSExcel::Range	range1,range2;
	m_range.AttachDispatch( m_worksheet.GetRows() , TRUE);
	range1.AttachDispatch( m_range.GetItem( COleVariant( (short)nFromRow ), vtMissing ).pdispVal , TRUE);
	range1.Select();
	range1.Copy(covOptional);
	range2.AttachDispatch(m_range.GetItem( COleVariant( (short)nToRow ), vtMissing ).pdispVal, TRUE);  
	range2.Select();
	m_worksheet.Paste(vtMissing,vtMissing);
	range1.ReleaseDispatch();
	range2.ReleaseDispatch();
	return TRUE;
}

BOOL MSExcelBooster::SetCellFont( const CString &sPlace1,const CString &sPlace2,COLORREF color,BOOL bBold,BOOL bItalic)
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)));
	MSExcel::Font ft;
	ft.AttachDispatch(m_range.GetFont());
	ft.SetColor(COleVariant((long)color));
	ft.SetBold(COleVariant((short)bBold));
	ft.SetItalic(COleVariant((short)bItalic));  
	ft.ReleaseDispatch();
	return TRUE;
}

long MSExcelBooster::GetUserRangeRowCount()
{
	MSExcel::Range	range,rowrange;
	range.AttachDispatch( m_worksheet.GetUsedRange() , TRUE);
	int		iRow;
	rowrange = m_range.GetRows();
	iRow = rowrange.GetCount();
	rowrange.ReleaseDispatch();
	range.ReleaseDispatch();
	return iRow;
}

long MSExcelBooster::GetUseRangeColumnCount()
{
	MSExcel::Range	range,colrange;
	range.AttachDispatch( m_worksheet.GetUsedRange() , TRUE);
	int		iColumn;
	colrange = m_range.GetColumns();
	iColumn = colrange.GetCount();
	colrange.ReleaseDispatch();
	range.ReleaseDispatch();
	return iColumn;
}

BOOL MSExcelBooster::SetCellVerAlign( const CString &sPlace1,const CString &sPlace2,const long nVerAlign )
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)));
	m_range.SetVerticalAlignment(COleVariant(nVerAlign));
	return TRUE;
}

BOOL MSExcelBooster::SetCellHorAlign( const CString &sPlace1,const CString &sPlace2,const long nHorAlign )
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)));
	m_range.SetHorizontalAlignment(COleVariant(nHorAlign));
	return TRUE;
}

MSExcel::Range MSExcelBooster::GetRange( const CString &sPlace1,const CString &sPlace2 )
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)));
	return m_range;
}

BOOL MSExcelBooster::SetCellBkgColor( const CString &sPlace1,const CString &sPlace2,COLORREF color )
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)));
	MSExcel::Interior it; 
	it.AttachDispatch(m_range.GetInterior()); 
	it.SetColor(COleVariant((long)color));
	return TRUE;
}

BOOL MSExcelBooster::SetCellBorder( const CString &sPlace1,const CString &sPlace2,int color,const short nWeight/*=2*/,const short nLineStyle/*=1*/ )
{
	m_range.AttachDispatch(m_worksheet.GetRange(COleVariant(sPlace1),COleVariant(sPlace2)));
	_variant_t  vRange1;   // 设置单元格的线；
	_variant_t  vRange2;
	_variant_t  vRange3;
	_variant_t  vRange4;
	// 线的样式：- no line; 1-solid; 2-big dot;3-small dot;4-dash dot; 5-dash dot dot;
	vRange1.vt = VT_I2;
	vRange1.lVal = nLineStyle; 
	// 线的粗细程度
	vRange2.vt = VT_I2;
	vRange2.lVal = nWeight;
	//线的颜色1-black;2-white;3-red;4-green;5-blue; 6-yellow; 7-pink;8-dark blue;
	vRange3.vt = VT_I2;
	vRange3.lVal = color;
	vRange4.vt = VT_I2;
	vRange4.lVal = RGB(0,0,0);
	m_range.BorderAround(vRange1, vRange2, vRange3, vRange4);
	return TRUE;
}

MSExcelBooster& endl( MSExcelBooster &excel )
{
	excel.SetCurrentRowIndex(excel.GetCurrentRowIndex()+1);
	excel.SetCurrentColumnIndex(1);
	return excel;
}

