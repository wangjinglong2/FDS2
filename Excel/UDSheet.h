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

	CUDRange    GetRange(int nRow, int nCol);//�������ǻ���1��ʼ
	CUDRange    GetRange(int nStartRow, int nEndRow, int nStartCol, int nEndCol);

	//��ȡ�������һ�������ݵ�������
	int         GetEndRow(int nCol);
 	 
	CUDColumns  GetColumns( int nStartCol, int nEndCol );
	CUDRows     GetRows(int nStartRow, int nEndRow);
 
	CUDPictures GetPictures();
	CUDShapes   GetShapes(); 
	
	//�����������ڵ�ǰλ��֮ǰ
	void        CopyInsterBefore();
	//�ƶ������λ��֮ǰ
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

	//����ͼƬ �о��������ô����
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
 
	//��������
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
	
	// ����ͼƬ width�� heightָ��Ϊ-1ʱ ����ͼƬԭ�гߴ�  δ��·����ȫ���
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

	//�����ݺ����
	void       SetLockAspectRatio(BOOL bLock=TRUE);
};

}