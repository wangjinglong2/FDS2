#pragma once
#include "EXCEL9.H"

class ExcelBoosterBase
{
public:
	ExcelBoosterBase(){};
	virtual ~ExcelBoosterBase(){};

public:
	//成员变量的存取函数
	CString get_FilePath(){
		return m_sFilePath;
	}
	void set_FilePath(const CString sFilePath){
		m_sFilePath = sFilePath;
	}
	BOOL get_Visible(){
		return m_bIsVisible;
	}
	void set_Visible(const BOOL bVisible){
		m_bIsVisible = bVisible;
	}
	CString get_FontStyle(){
		return m_sFontStyle;
	}
	void set_FontStyle(const CString sFontStyle){
		m_sFontStyle = sFontStyle;
	}
	int get_FontSize(){
		return m_nFontSize;
	}
	void set_FontSize(const int nFontSize){
		m_nFontSize = nFontSize;
	}
	int get_nTotalSheet(){
		return m_nTotalSheet;
	}
	void set_nTotalSheet(const int nTotalSheet){
		m_nTotalSheet = nTotalSheet;
	}
	void set_QuitFlag(const bool bQuitFlag){
		m_bQuitExcel = bQuitFlag;
	}
	void SetCurrentColumnIndex(const long value){
		m_nCurrentColumnIndex = value;
	}
	long GetCurrentColumnIndex(){
		return m_nCurrentColumnIndex;
	}
	void SetCurrentRowIndex(const long value){
		m_nCurrentRowIndex = value;
	}
	long GetCurrentRowIndex()
	{
		return m_nCurrentRowIndex;
	}
	void SetStartColumnIndex(const long value){
		m_nStartColumnIndex = value;
	}
	long GetStartColumnIndex(){
		return m_nStartColumnIndex;
	}
	void SetDisplayAlerts(BOOL bDisplayAlerts){
		m_bDisplayAlerts = bDisplayAlerts;
	}

protected:
	CString m_sFilePath;
	BOOL m_bIsVisible;
	BOOL m_bDisplayAlerts;
	CString m_sFontStyle;
	int m_nFontSize;
	int m_nTotalSheet;
	bool m_bInitComTrue;
	bool m_bQuitExcel;
	long m_nStartColumnIndex;
	long m_nCurrentColumnIndex;
	long m_nCurrentRowIndex;
};

class MSExcelBooster : public ExcelBoosterBase
{
public:
	//	将光标转到指定表格单元
	BOOL MoveTo(const long nRowIndex,const long nColunIndex);
	//	转到下一列表格单元
	void NextColumn();

	MSExcelBooster &operator<<(const int value);
	MSExcelBooster &operator<<(const long value);
	MSExcelBooster &operator<<(const short value);
	MSExcelBooster &operator<<(const double value);
	MSExcelBooster &operator<<(const TCHAR *sValue);
	MSExcelBooster &operator<<( const CString& sValue);
	MSExcelBooster &operator>>(int &value);
	MSExcelBooster &operator>>(long &value);
	MSExcelBooster &operator>>(short &value);
	MSExcelBooster &operator>>(double &value);
	MSExcelBooster &operator>>(TCHAR *sValue);
	MSExcelBooster &operator>>(CString &sValue);

	MSExcelBooster& operator << (MSExcelBooster& (*op)(MSExcelBooster&)) {return (*op)(*this);}
	MSExcelBooster& operator >> (MSExcelBooster& (*op)(MSExcelBooster&)) {return (*op)(*this);}
	friend MSExcelBooster& endl(MSExcelBooster &excel);

public:
	//构造函数和解析函数
	MSExcelBooster();
	virtual ~MSExcelBooster();

	//	返回值：-1：没有初始化COM!; 0：没有安装Excel2000!; 1：初始化成功。
	BOOL InitExcelCOM();
	BOOL Quit();
	////	释放COM对象
	//void ReleaseExcelCom();
	BOOL OpenExcelBook(CString sFileName);
	//void OpenOneWorkSheets(CString sFileName);
	BOOL NewExcelBook();
	BOOL OpenExcelApp();

	//	设定当前表页
	BOOL SetCurWorkSheet(int nCurSheet);
	//设定当前表页，并命名
	BOOL SetCurWorkSheet(int nCurSheet,const TCHAR *sSheetName);
	//设定指定名称表页为当前表页
	BOOL SetCurWorkSheet(const TCHAR *sSheetName);
	//	删除指定表页
	BOOL DeleteSheet(const short nItem);
	//创建一个新的表页
	BOOL AddSheet();

	BOOL SaveExcel();
	BOOL SaveAsExcel(CString sFileName);
	BOOL SetCellValue(int row, int col,CString sValue,int Align=1);
	CString GetCellValue(int row, int col);
	BOOL SetRowHeight(int row, CString height);
	CString GetRowHeight(int row);
	BOOL SetColumnWidth(int col,CString width);
	CString GetColumnWidth(int col);
	//根据行列指获取单元格位置
	CString GetCellPos( int row, int col );
	//得到内容为VAR的单元格地址
	CString GetCellPos(const CString& var);
	BOOL GetCellPos(const CString& var,int& iRow,int& iColumn);

	//合并单元格
	BOOL MergeCell(const CString& sPlace1,const CString& sPlace2);
	//填充合并单元格内容
	BOOL FillMerge(const CString &place1,const CString &place2,const CString &sText);
	//插入图片
	BOOL InsertPicture(const CString& sPlace1,const CString& sPlace2,const CString& sPicPath,BOOL bZoom,const int nGap = 0);
public:
	//	删除指定行
	BOOL DeleteRow(int nRow);
	BOOL DeleteRows( int nStartRow,int nEndRow );
	BOOL InsertRow(int nStartRow,int nRowCnt);
	BOOL CopyRow(int nFromRow,int nToRow);

	//	得到表格单元的颜色，返回为颜色索引
	int GetCellColorIndex(const CString &sPlace);
	BOOL SetCellColorIndex(const CString &sPlace,int colorindex);
	//	得到行数
	long GetUserRangeRowCount();
	//	得到列数
	long GetUseRangeColumnCount();
//
//	//	得到表格水平垂直方式
//	//	返回或设置指定对象的垂直对齐方式。可为下列 XlVAlign 常量之一： 
//	//	xlVAlignBottom、xlVAlignCenter、xlVAlignDistributed、xlVAlignJustify 
//	//	或 xlVAlignTop。Long 类型，可读写。
//
//	//	垂直队齐<缺省为2>：1--靠上，2--居中，3--靠下，4--两端对齐，5--分散对齐
//	long GetCellVerAlign(const CString &sPlace);
//	void SetCellVerAlign(const CString &sPlace,const long nVerAlign);
//
//	//	得到表格水平对齐方式
//	//	返回或设置指定对象的水平对齐方式。对所有对象，可为以下 XlHAlign 常数之一：
//	//	xlHAlignCenter、xlHAlignDistributed、xlHAlignJustify、xlHAlignLeft 或者 xlHAlignRight。
//	//	另外，对于 Range 或 Style 对象，可以将此属性设置为 xlHAlignCenterAcrossSelection、xlHAlignFill 或 
//	//	xlHAlignGeneral。Variant 类型，可读写。
//
//	//	水平对齐<缺省为1>：1--常规，2--靠左，3--居中，4--靠右，5--填充；
//	//	6--两端对齐；7--跨列队齐；8--分散对齐
//	long GetCellHorAlign(const CString &sPlace);
//	void SetCellHorAlign(const CString &sPlace,const long nHorAlign);
//
//	//	得到指定表的名称
//	void GetSheetName(const short nIndex,CString &sName);
//
//	//	得到行数
//	long GetUserRangeRowCount();
//
//	//	得到列数
//	long GetUseRangeColumnCount();
//
//	//	得到表格单元内容(只能获得文本的内容，不能获得如OLE或者图片等信息内容)
//	void GetCellValue(CString sPlace, CString &value);
//
//	//	初始化页面设置
//	void InitPageSetup(double dTopMargin,double BottomMargin,double dLeftMargin,double dRightMargin,
//		double dHeaderMargin,double dFooterMargin);
//	//	绘制表格边线
//	void DrawBorder(CString place1,CString place2,const short nWeight=2,const short nLineStyle=1);
//
//	//	nBorderIndex
//	//	5--DiagonalDown
//	//	6--DiagonalUp
//	//	7--EdgeLeft
//	//	8--EdgeTop
//	//	9--EdgeBottom
//	//	10--EdgeRight
//	//	11--InsideVertical
//	//	12--InsideHorizontal
//	void DrawBorderLine(CString place1,CString place2,const long nBorderIndex=4,const long nWeight=2,
//		const long nLineStyle=1,const long nColorIndex=1);
//
//	//	插入图片到Excel表中(支持Jpg和bmp格式图片)
//	void InsertPicture(const TCHAR *sPlace1,const TCHAR *sPlace2,const TCHAR *sPicPath,const int nGap = 0);
//
//	//	往表格单元填写数据
//	BOOL FillCell(const long nRow/*>=1*/,const int nCol/*>=1*/,const TCHAR *sText);
//	void FillCell(const TCHAR *sPlace,const TCHAR *sText);
//	void FillCell(const TCHAR *sPlace,const TCHAR *sText,const long lBkColor);
//	void FillCell(const CString &sPlace,const CString &sText,const long nHorAlign,
//		const long nVerAlign,const long nColor);
//	//	填充合并表格单元数据
//	void FillMerge(const TCHAR *place1,const TCHAR *place2,const TCHAR *sText);
//	void FillMerge(CString &place1,CString &place2,CString &sText,const double dUserHeight,const short nHorAlign=1,
//		const short nVerAlign=2,const double dOrientation=.0,const TCHAR *sFontStyle=_T(""));
//	void FillMerge(const CString &place1,const CString &place2,const CString &sText,
//		const long nHorAlign,const long nVerAlign,const long nColor);
//
//	//	设定表格列宽
//	BOOL SetColumnWidth(int nCol,double dShellWidth=20);
//
//	//	得到表格列宽
//	double GetColumnWidth(int nCol);
//
//	//	设定表格行高
//	BOOL SetRowHeight(int nRow,double dShellHeight=20);
//
//	//得到表格行高
//	double GetRowHeight(int nRow);
//
//	//打印当前表页所有内容
//	void PrintOut(const UINT from,const UINT to,const UINT copies=1);
//
//	//设置打印区域:eg: _T("$A$1:$I$47")
//	void SetPrintArea(LPCTSTR lpszNewValue);
//	//设置和获取当前表页可见性:xlSheetHidden = 0,xlSheetVeryHidden = 2,xlSheetVisible = -1
//	long GetSheetVisible();
//	void SetSheetVisible(long nNewValue);
//	//在当前表页插入一列：在指定列后插入一列
//	BOOL InsertColumn(const long nColumnIdx,const BOOL bXlToRight);
//
//	void set_Visible(const bool bVisible);
//	long GetSheetCount();
//
//public:	//CExcelBoosterEx函数
//	BOOL Save();
//	BOOL SaveAs(CString szFileName);
//	BOOL SaveAs(CString szFileName,const long xlFileFormat);
//	void Quit();

//	long GetUsedRangeRow();
//	CString GetCurCellPos();
//	CString GetCellPos(int nRowIndex,int nColIndex);
//
//	void SetHidden(int nRowIndex,BOOL bNewValue);
//
//	//返回单元格是否隐藏
//	BOOL GetHidden(const CString &szCell);
//
//	//得到内容为VAR的单元格地址
//	CString GetCell(const CString& var);
//	//查找替换
//	void FindAndReplace(const CString &strFind, const CString &strReplace);
//
//	BOOL GetCell(const CString& var,int& nRow,int& nCom);
//	//得到单元格的内容
//	CString GetItem(const CString& var);
//
//	void CloseWorkbook();
//
//	BOOL CopySheet(const short nIndexFrom,const short nIndexTo);
//	//	设置指定表的名称
//	void SetSheetName(const short nIndex, const CString &sName);
//	long GetCurSheetIndex();
//	void Visible(BOOL bVisible);
//	//szCell单元格位置,
//	//szFormat单元格的格式
//	void SetNumberFormatLocal(CString szCell,CString szFormat);
//	void RangeMerge(CString szCell1,CString szCell2);
//	void CopyRange(const int nSheetItem,CString szCell);
//	BOOL CopyRefersToRange(const CString &sRefRange,const CString &sToRange);
//
//	BOOL InsertSheet(CString szXlsFileName);
//	long GetUsedRangeCol();
//
//	void SetColor(const CString& strFrom,const CString& strTo,int nColor,int nRow);
//
//	BOOL AddPicture(const CString& strFileName,const CString& strCell1,const CString& strCell2);
//	//--导出Excel中的图片到外部位图文件 pStrCellName:单元格名称，pStrBmpSavePath：位图保存路径
//	BOOL ExportPicToBmpFile(const TCHAR *pStrCellName, const TCHAR *pStrBmpSavePath);
//	BOOL SaveToTxt(int nRow,int nCom,int nRowCount,int nComCount,const CString& strHeader = _T(""));
//	BOOL SaveToArray(std::vector< std::vector<std::string> >& data);
//	BOOL GetDisplayAlerts();
//	void SetDisplayAlerts(BOOL bNewValue);
//	void SetFontColor(const CString& strFrom,const CString& strTo,int nColor,int nRow);
//	BOOL ToRange(int nRow,int nCol,CStringArray& dataArr);
//	BOOL ToRange(CStringArray& dataArr);
//
//	BOOL ToRange(int nRow, int nCol, const std::vector<std::vector<COleVariant> > &vals);
//	BOOL ToRange(const std::vector<std::vector<COleVariant> > &vals);
//
//	BOOL GetUsedRangeValue(std::vector<std::vector<COleVariant> > &vals);
//	BOOL SetUsedRangeValue(const std::vector<std::vector<COleVariant> > &vals);
//
//	//自适应行高
//	void AutoFitRowHeight(int nItem);
//	CString GetCellValueByName(const CString& strName);
//
//	//自适应列宽:eg,szCol=_T("A") or szCol=_T("AB")
//	void AuotFitColWidth(CString szCol);
//	long GetWorkSheetPrintPages();
//	long GetWorkBookPrintPages();
//	void SetCenterHeader(const CString &header);
//	void SetRightHeader(const CString &header);
//	void SetHeaderMargin(double margin);
//	BOOL DeleteSheet(const CString &sheetName);
//	BOOL PrintWorkBook();

protected:

	MSExcel::_Application m_application;    
	MSExcel::Workbooks m_workbooks;   
	MSExcel::_Workbook m_workbook;    
	MSExcel::Worksheets m_worksheets;   
	MSExcel::_Worksheet m_worksheet; 
	MSExcel::Range m_range;  
	MSExcel::Range m_cell;  
	MSExcel::Font m_font;  

//public:
//	static CString g_OleXlsApp;
//	BOOL CopySheet(MSExcel::_Worksheet& worksheetFrom,MSExcel::_Worksheet& worksheetTo);
//	void CopyRange(MSExcel::Range& UsedRange,CString szCell);
//	BOOL GetName(const CString& strNameVal,MSExcel::Range& reRange);
//	BOOL GetNameWorkSheet(CString szName, MSExcel::Range& reRange);
};
