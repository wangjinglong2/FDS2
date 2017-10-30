#pragma once
#include "furniture.h"

class Board:public FurniturePart
{
	DECLARE_DYNAMIC(Board)

	static int g_nVersion;//版本号
public:
	Board(void);
	virtual ~Board(void);

	virtual BOOL Rebulid();
	virtual BOOL Clone(Board* pBoard);
public:
	//edit and query function
	void SetPosition(const AcGePoint3d& ptOrg);
	AcGePoint3d GetPosition();
	void SetDirection(const AcGeVector3d& vecx,const AcGeVector3d& vecy,const AcGeVector3d& vecz);
	void GetDirection(AcGeVector3d& vecx,AcGeVector3d& vecy,AcGeVector3d& vecz);
	void SetDimension(double dLen,double dWidth,double dThick);
	void GetDimension(double& dLen,double& dWidth,double& dThick);
	void GetFaceCoordSystem(int iFaceNo,AcGePoint3d& ptOrg,AcGePoint3d& ptX,AcGePoint3d& ptY,AcGePoint3d& ptZ);
	void GetFaceDimension(int iFaceNo,double& dLen,double& dWidth);
	BOOL IsEmpty();
	BOOL TransformBy(AcGeMatrix3d mat);
	BOOL Highlight(BOOL bHighlight);
	BOOL GetBoardExtent(AcDbExtents& extent);
	AcDbObjectIdArray GetHardware();
public:
	virtual void GiDrawFace(int iFaceNo,int color);
	virtual void GetFaceNo(int& iFaceNo);
protected:
	virtual struct resbuf* PrepareXData();
	virtual BOOL GetObjectXData( AcDbObjectId idPart,struct resbuf*& pBuf);
	virtual BOOL PrePareObject();
protected:
	AcGePoint3d		m_ptOrg;
	AcGeVector3d	m_vecX,m_vecY,m_vecZ;
	double			m_dLength,m_dWidth,m_dThick;
	AcDbObjectId	m_idBaseBd;
	CArray<FdmCoordSystem,FdmCoordSystem&> m_pFixFaceArray;	//定位表面集
	AcDbObjectIdArray	m_idHardwareArray;
};

