#pragma once
#include "EXCEL9.H"

class ExcelBoosterBase
{
public:
	ExcelBoosterBase(){};
	virtual ~ExcelBoosterBase(){};

public:
	//��Ա�����Ĵ�ȡ����
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
	//	�����ת��ָ�����Ԫ
	BOOL MoveTo(const long nRowIndex,const long nColunIndex);
	//	ת����һ�б��Ԫ
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
	//���캯���ͽ�������
	MSExcelBooster();
	virtual ~MSExcelBooster();

	//	����ֵ��-1��û�г�ʼ��COM!; 0��û�а�װExcel2000!; 1����ʼ���ɹ���
	BOOL InitExcelCOM();
	BOOL Quit();
	////	�ͷ�COM����
	//void ReleaseExcelCom();
	BOOL OpenExcelBook(CString sFileName);
	//void OpenOneWorkSheets(CString sFileName);
	BOOL NewExcelBook();
	BOOL OpenExcelApp();

	//	�趨��ǰ��ҳ
	BOOL SetCurWorkSheet(int nCurSheet);
	//�趨��ǰ��ҳ��������
	BOOL SetCurWorkSheet(int nCurSheet,const TCHAR *sSheetName);
	//�趨ָ�����Ʊ�ҳΪ��ǰ��ҳ
	BOOL SetCurWorkSheet(const TCHAR *sSheetName);
	//	ɾ��ָ����ҳ
	BOOL DeleteSheet(const short nItem);
	//����һ���µı�ҳ
	BOOL AddSheet();

	BOOL SaveExcel();
	BOOL SaveAsExcel(CString sFileName);
	BOOL SetCellValue(int row, int col,CString sValue,int Align=1);
	CString GetCellValue(int row, int col);
	BOOL SetRowHeight(int row, CString height);
	CString GetRowHeight(int row);
	BOOL SetColumnWidth(int col,CString width);
	CString GetColumnWidth(int col);
	//��������ָ��ȡ��Ԫ��λ��
	CString GetCellPos( int row, int col );
	//�õ�����ΪVAR�ĵ�Ԫ���ַ
	CString GetCellPos(const CString& var);
	BOOL GetCellPos(const CString& var,int& iRow,int& iColumn);

	//�ϲ���Ԫ��
	BOOL MergeCell(const CString& sPlace1,const CString& sPlace2);
	//���ϲ���Ԫ������
	BOOL FillMerge(const CString &place1,const CString &place2,const CString &sText);
	//����ͼƬ
	BOOL InsertPicture(const CString& sPlace1,const CString& sPlace2,const CString& sPicPath,BOOL bZoom,const int nGap = 0);
public:
	//	ɾ��ָ����
	BOOL DeleteRow(int nRow);
	BOOL DeleteRows( int nStartRow,int nEndRow );
	BOOL InsertRow(int nStartRow,int nRowCnt);
	BOOL CopyRow(int nFromRow,int nToRow);

	//	�õ����Ԫ����ɫ������Ϊ��ɫ����
	int GetCellColorIndex(const CString &sPlace);
	BOOL SetCellColorIndex(const CString &sPlace,int colorindex);
	//	�õ�����
	long GetUserRangeRowCount();
	//	�õ�����
	long GetUseRangeColumnCount();
//
//	//	�õ����ˮƽ��ֱ��ʽ
//	//	���ػ�����ָ������Ĵ�ֱ���뷽ʽ����Ϊ���� XlVAlign ����֮һ�� 
//	//	xlVAlignBottom��xlVAlignCenter��xlVAlignDistributed��xlVAlignJustify 
//	//	�� xlVAlignTop��Long ���ͣ��ɶ�д��
//
//	//	��ֱ����<ȱʡΪ2>��1--���ϣ�2--���У�3--���£�4--���˶��룬5--��ɢ����
//	long GetCellVerAlign(const CString &sPlace);
//	void SetCellVerAlign(const CString &sPlace,const long nVerAlign);
//
//	//	�õ����ˮƽ���뷽ʽ
//	//	���ػ�����ָ�������ˮƽ���뷽ʽ�������ж��󣬿�Ϊ���� XlHAlign ����֮һ��
//	//	xlHAlignCenter��xlHAlignDistributed��xlHAlignJustify��xlHAlignLeft ���� xlHAlignRight��
//	//	���⣬���� Range �� Style ���󣬿��Խ�����������Ϊ xlHAlignCenterAcrossSelection��xlHAlignFill �� 
//	//	xlHAlignGeneral��Variant ���ͣ��ɶ�д��
//
//	//	ˮƽ����<ȱʡΪ1>��1--���棬2--����3--���У�4--���ң�5--��䣻
//	//	6--���˶��룻7--���ж��룻8--��ɢ����
//	long GetCellHorAlign(const CString &sPlace);
//	void SetCellHorAlign(const CString &sPlace,const long nHorAlign);
//
//	//	�õ�ָ���������
//	void GetSheetName(const short nIndex,CString &sName);
//
//	//	�õ�����
//	long GetUserRangeRowCount();
//
//	//	�õ�����
//	long GetUseRangeColumnCount();
//
//	//	�õ����Ԫ����(ֻ�ܻ���ı������ݣ����ܻ����OLE����ͼƬ����Ϣ����)
//	void GetCellValue(CString sPlace, CString &value);
//
//	//	��ʼ��ҳ������
//	void InitPageSetup(double dTopMargin,double BottomMargin,double dLeftMargin,double dRightMargin,
//		double dHeaderMargin,double dFooterMargin);
//	//	���Ʊ�����
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
//	//	����ͼƬ��Excel����(֧��Jpg��bmp��ʽͼƬ)
//	void InsertPicture(const TCHAR *sPlace1,const TCHAR *sPlace2,const TCHAR *sPicPath,const int nGap = 0);
//
//	//	�����Ԫ��д����
//	BOOL FillCell(const long nRow/*>=1*/,const int nCol/*>=1*/,const TCHAR *sText);
//	void FillCell(const TCHAR *sPlace,const TCHAR *sText);
//	void FillCell(const TCHAR *sPlace,const TCHAR *sText,const long lBkColor);
//	void FillCell(const CString &sPlace,const CString &sText,const long nHorAlign,
//		const long nVerAlign,const long nColor);
//	//	���ϲ����Ԫ����
//	void FillMerge(const TCHAR *place1,const TCHAR *place2,const TCHAR *sText);
//	void FillMerge(CString &place1,CString &place2,CString &sText,const double dUserHeight,const short nHorAlign=1,
//		const short nVerAlign=2,const double dOrientation=.0,const TCHAR *sFontStyle=_T(""));
//	void FillMerge(const CString &place1,const CString &place2,const CString &sText,
//		const long nHorAlign,const long nVerAlign,const long nColor);
//
//	//	�趨����п�
//	BOOL SetColumnWidth(int nCol,double dShellWidth=20);
//
//	//	�õ�����п�
//	double GetColumnWidth(int nCol);
//
//	//	�趨����и�
//	BOOL SetRowHeight(int nRow,double dShellHeight=20);
//
//	//�õ�����и�
//	double GetRowHeight(int nRow);
//
//	//��ӡ��ǰ��ҳ��������
//	void PrintOut(const UINT from,const UINT to,const UINT copies=1);
//
//	//���ô�ӡ����:eg: _T("$A$1:$I$47")
//	void SetPrintArea(LPCTSTR lpszNewValue);
//	//���úͻ�ȡ��ǰ��ҳ�ɼ���:xlSheetHidden = 0,xlSheetVeryHidden = 2,xlSheetVisible = -1
//	long GetSheetVisible();
//	void SetSheetVisible(long nNewValue);
//	//�ڵ�ǰ��ҳ����һ�У���ָ���к����һ��
//	BOOL InsertColumn(const long nColumnIdx,const BOOL bXlToRight);
//
//	void set_Visible(const bool bVisible);
//	long GetSheetCount();
//
//public:	//CExcelBoosterEx����
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
//	//���ص�Ԫ���Ƿ�����
//	BOOL GetHidden(const CString &szCell);
//
//	//�õ�����ΪVAR�ĵ�Ԫ���ַ
//	CString GetCell(const CString& var);
//	//�����滻
//	void FindAndReplace(const CString &strFind, const CString &strReplace);
//
//	BOOL GetCell(const CString& var,int& nRow,int& nCom);
//	//�õ���Ԫ�������
//	CString GetItem(const CString& var);
//
//	void CloseWorkbook();
//
//	BOOL CopySheet(const short nIndexFrom,const short nIndexTo);
//	//	����ָ���������
//	void SetSheetName(const short nIndex, const CString &sName);
//	long GetCurSheetIndex();
//	void Visible(BOOL bVisible);
//	//szCell��Ԫ��λ��,
//	//szFormat��Ԫ��ĸ�ʽ
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
//	//--����Excel�е�ͼƬ���ⲿλͼ�ļ� pStrCellName:��Ԫ�����ƣ�pStrBmpSavePath��λͼ����·��
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
//	//����Ӧ�и�
//	void AutoFitRowHeight(int nItem);
//	CString GetCellValueByName(const CString& strName);
//
//	//����Ӧ�п�:eg,szCol=_T("A") or szCol=_T("AB")
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
