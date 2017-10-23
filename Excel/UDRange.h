#pragma once
#include "UDIDispath.h"
#include<string>

using namespace std;

namespace UDExcel{
#define  TEXT_H_ALIGN_DEFAULT 1
#define  TEXT_H_ALIGN_LEFT    -4108
#define  TEXT_H_ALIGN_CENTER  -4131
#define  TEXT_H_ALIGN_RIGHT   -4152

#define  TEXT_V_ALIGN_DEFAULT 2
#define  TEXT_V_ALIGN_TOP     -4108
#define  TEXT_V_ALIGN_MIDDLE  -4160
#define  TEXT_V_ALIGN_BOTTOM  -4107

class CUDFont;
class CUDRange : public CUDIDispath
{
public: 
	CUDRange(CUDRange& range){ this->m_lpDispatch = range.m_lpDispatch; }
	CUDRange(IDispatch*pRange){ this->m_lpDispatch = pRange; }
	CUDRange& operator = (CUDRange& Range){ this->m_lpDispatch = Range.m_lpDispatch; return *this; }
	virtual ~CUDRange(){}
	
	
	BOOL        IsNull(){ return m_lpDispatch==0; }

	//=SUM(A16:A65536)求平均  可以设置公式的
	void        SetValue(string sText);

	string      GetValue();
	void        Merge();
	void        UnMerge();
	void        SetBKColor(int nColor);

	void        SetTextHAlign(int nAlign=TEXT_H_ALIGN_DEFAULT);
	void        SetTextVAlign(int nAlign=TEXT_V_ALIGN_DEFAULT);
	void        SetSelect();

	CUDFont     GetFont();
	double      Top();
	double      Left();
	double      Height();
	double      Width();

	void        ClearAll();//清空所有
	void        ClearContents();//清空内容
	void        ClearFormats();//清空格式
	void        ClearComments();//清空批注

	//          获取某一列的最后一行
	int         GetEndRow();
};

class CUDFont : public CUDIDispath
{
public:
	CUDFont(CUDFont& ft){ this->m_lpDispatch = ft.m_lpDispatch; }
	CUDFont(IDispatch*pFt){ this->m_lpDispatch = pFt; }
	CUDFont& operator = (CUDFont& ft){ this->m_lpDispatch = ft.m_lpDispatch; return *this; }
	virtual ~CUDFont(){}

	void         SetColor(int nColor);
	void         SetSize(int nSize);
	void         SetName(string sName);
	void         SetItalic(BOOL bItalic=TRUE);
	void         SetBold(BOOL bBold=TRUE);
};
}