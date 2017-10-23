#include "UDRange.h"
#include "UDIDispath.h"
#pragma once
 namespace UDExcel{

class CUDColumns;
class CUDRows;
class CUDPictures;
class CUDShapes;
class CUDPicture;
class CUDShapeRange;

class CUDSheet:public CUDIDispath
{
public:
	CUDSheet(CUDSheet& Sheet){ this->m_lpDispatch = Sheet.m_lpDispatch; }
	CUDSheet(IDispatch*pSheet){ this->m_lpDispatch = pSheet; }
	CUDSheet& operator = (CUDSheet& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDSheet(){ /*if(!IsNull())m_pDocs->Release();*/ }
	
	 
	void        SetName(string sName);
	string      GetName();
	void        Delete();
	void        SetVisale( BOOL bVisable=true);
	void        SetActive();

	CUDRange    GetRange(int nRow, int nCol);//参数都是基于1开始
	CUDRange    GetRange(int nStartRow, int nEndRow, int nStartCol, int nEndCol);

	//获取该列最后一个有数据的行索引
	int         GetEndRow(int nCol);
 	 
	CUDColumns  GetColumns( int nStartCol, int nEndCol );
	CUDRows     GetRows(int nStartRow, int nEndRow);
 
	CUDPictures GetPictures();
	CUDShapes   GetShapes(); 
	
	//拷贝并插入在当前位置之前
	void        CopyInsterBefore();
	//移动到这个位置之前
	void        MoveBefore(CUDSheet&sheet);
};


class CUDColumns:public CUDIDispath
{
public:
	IDispatch*  m_pSheet;
	CUDColumns(CUDColumns& Sheet){ this->m_pSheet = Sheet.m_pSheet; }
	CUDColumns(IDispatch*pSheet){ this->m_pSheet = pSheet; }
	CUDColumns& operator = (CUDColumns& Docs){ this->m_pSheet = Docs.m_pSheet; return *this; }
	virtual ~CUDColumns(){ /*if(!IsNull())m_pDocs->Release();*/ }

	void     SetWidth(double dWidth);
	void     SetVisale( BOOL bVisable=true);
	void     Delete();
	void     SetSelect();
	
};

class CUDRows:public CUDIDispath
{
public: 
	CUDRows(CUDRows& Sheet){ this->m_lpDispatch = Sheet.m_lpDispatch; }
	CUDRows(IDispatch*pSheet){ this->m_lpDispatch = pSheet; }
	CUDRows& operator = (CUDRows& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDRows(){ /*if(!IsNull())m_pDocs->Release();*/ }
	
	void     SetHeight(double dWidth);
	void     Delete();
	void     SetVisale( BOOL bVisable=true);
	void     SetSelect();
};


class CUDPictures:public CUDIDispath
{
public: 
	CUDPictures(CUDPictures& Sheet){ this->m_lpDispatch = Sheet.m_lpDispatch; }
	CUDPictures(IDispatch*pSheet){ this->m_lpDispatch = pSheet; }
	CUDPictures& operator = (CUDPictures& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDPictures(){ /*if(!IsNull())m_pDocs->Release();*/ }

	//插入图片 感觉这个不怎么靠谱
	CUDPicture InsertPicture(string sPath);	
	int        GetCount();
	CUDPicture GetItem( int nIndex );

};

class CUDPicture:public CUDIDispath
{
public: 
	CUDPicture(CUDPicture& Sheet){ this->m_lpDispatch = Sheet.m_lpDispatch; }
	CUDPicture(IDispatch*pSheet){ this->m_lpDispatch = pSheet; }
	CUDPicture& operator = (CUDPicture& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDPicture(){ /*if(!IsNull())m_pDocs->Release();*/ }
	
	
	void    SetWidth(double dWidth);
	void    SetHeight(double dHeight);
 
	//设置名称
	void    SetName(string sName);
	string  GetName();
	void    Delete();
	CUDShapeRange GetShapeRange();
};


class CUDShapes:public CUDIDispath
{
public: 
	CUDShapes(CUDShapes& Sheet){ this->m_lpDispatch = Sheet.m_lpDispatch; }
	CUDShapes(IDispatch*pSheet){ this->m_lpDispatch = pSheet; }
	CUDShapes& operator = (CUDShapes& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDShapes(){ /*if(!IsNull())m_pDocs->Release();*/ }
	
	// 插入图片 width和 height指定为-1时 保持图片原有尺寸  未做路径安全检查
	CUDPicture      AddPicture(string sPath, int Left, int Top, int Width, int Height);
  	int             GetCount();
	CUDShapeRange   GetRange( int nIndex );
};

class CUDShapeRange:public CUDIDispath
{
public: 
	CUDShapeRange(CUDShapeRange& Sheet){ this->m_lpDispatch = Sheet.m_lpDispatch; }
	CUDShapeRange(IDispatch*pSheet){ this->m_lpDispatch = pSheet; }
	CUDShapeRange& operator = (CUDShapeRange& Docs){ this->m_lpDispatch = Docs.m_lpDispatch; return *this; }
	virtual ~CUDShapeRange(){ /*if(!IsNull())m_pDocs->Release();*/ }

	//设置纵横比例
	void       SetLockAspectRatio(BOOL bLock=TRUE);
};

}