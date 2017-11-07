#pragma once
#include "Furniture.h"
#include "Board.h"
#include "hardwaredata.h"


class HardWare:public FurniturePart
{
	DECLARE_DYNAMIC(HardWare)
	static int	g_nVersion;
public:
	HardWare();
	virtual ~HardWare();
	virtual BOOL Rebulid()=0;
	virtual void SetPosition(AcGePoint3d ptOrg,AcGePoint3d ptX,AcGePoint3d ptY,AcGePoint3d ptZ);
	void SetFixBdArray(AcDbObjectIdArray idFixBdArray){m_idFixBdArray = idFixBdArray;};
	AcDbObjectIdArray GetFixBdArray(){return m_idFixBdArray;};

protected:
	virtual struct resbuf* PrepareXData();
	virtual BOOL GetObjectXData(AcDbObjectId idPart,struct resbuf *&xdata);
	virtual BOOL PrePareObject();
	virtual BOOL GetHoles(CArray<Hole,Hole&>& holes){return TRUE;};
	virtual BOOL GetBoardHoles(AcDbObjectId idBoard,CArray<Hole,Hole&>& holes){return TRUE;};
protected:
	AcGePoint3d m_ptBase,m_ptX,m_ptY,m_ptZ;	
	BOOL m_bCallDwgBlock;
	CString m_sDwgName;	
	AcDbObjectIdArray	m_idFixBdArray;
	CArray<Hole,Hole&>	m_holes;	//所有的孔位数据

};

class BiasConnecter:public HardWare
{
	DECLARE_SERIAL(BiasConnecter)
	static int	g_nVersion;
public:
	BiasConnecter(void);
	~BiasConnecter(void);
	BOOL PrepareHwData(Board& firstBd,Board& secondBd,BOOL bPosFace,double dLayDist);
	virtual BOOL Rebulid();
	virtual BOOL GetHoles(CArray<Hole,Hole&>& holes);
	virtual BOOL GetBoardHoles(AcDbObjectId idBoard,CArray<Hole,Hole&>& holes);
protected:
	BOOL getDwgPath(CString& sDwgPath);
	BOOL HasBlock(CString szRecordName,AcDbObjectId & id);
	Acad::ErrorStatus addBlkRecordFromDWG(CString szDwgPathName,AcDbObjectId& idBlockRecord,CString szBlkName = _T(""));
	AcGeMatrix3d	m_matrix;

protected:
	virtual struct resbuf* PrepareXData();
	virtual BOOL GetObjectXData(AcDbObjectId idPart,struct resbuf *&xdata);
	virtual BOOL PrePareObject();
private:
	BiasConnecterData	m_biasData;
};

