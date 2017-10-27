#pragma once
#include "EXCEL9.H"

class ExcelBoosterBase
{
public:
	ExcelBoosterBase(){};
	virtual ~ExcelBoosterBase(){};

public:
	//单元格对齐方式
	enum AlignType
	{
		HORCENALIGN = -4108,
		HORLEFTALIGN = -4131,
		HORRIGHTALIGN = -4152,
		VERCENALIGN = -4108,
		VERLEFTALIGN = -4160,
		VERRIGHTALITN = -4107
	};
	//单元格移动方向
	enum XlDirection
	{
		xlDown        = -4121,
		xlToLeft    = -4159,
		xlToRight    = -4161,
		xlUp        = -4162
	};

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

	BOOL InitExcelCOM();
	BOOL Quit();
	BOOL OpenExcelBook(CString sFileName);
	BOOL NewExcelBook();
	BOOL OpenExcelApp();
	BOOL SaveExcel();
	BOOL SaveAsExcel(CString sFileName);

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

	//	删除指定行
	BOOL DeleteRow(int nRow);
	BOOL DeleteRows( int nStartRow,int nEndRow );
	BOOL InsertRow(int nStartRow,int nRowCnt);
	BOOL CopyRow(int nFromRow,int nToRow);

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
	//	设置单元格字体
	BOOL SetCellFont( const CString &sPlace1,const CString &sPlace2,COLORREF color,BOOL bBold = FALSE,BOOL bItalic = FALSE);
	//设置垂直对齐
	BOOL SetCellVerAlign(const CString &sPlace1,const CString &sPlace2,const long nVerAlign);
	//设置水平对齐
	BOOL SetCellHorAlign(const CString &sPlace1,const CString &sPlace2,const long nHorAlign);
	//设置单元格背景色
	BOOL SetCellBkgColor(const CString &sPlace1,const CString &sPlace2,COLORREF color);
	//设置单元格边线
	BOOL SetCellBorder(const CString &sPlace1,const CString &sPlace2,int color,const short nWeight=2,const short nLineStyle=1);
public:
	//得到行数
	long GetUserRangeRowCount();
	//得到列数
	long GetUseRangeColumnCount();
	//获取单元格区域
	MSExcel::Range GetRange(const CString &sPlace1,const CString &sPlace2);

protected:
	MSExcel::_Application m_application;    
	MSExcel::Workbooks m_workbooks;   
	MSExcel::_Workbook m_workbook;    
	MSExcel::Worksheets m_worksheets;   
	MSExcel::_Worksheet m_worksheet; 
	MSExcel::Range m_range;  
	MSExcel::Range m_cell;  
	MSExcel::Font m_font; 
};
