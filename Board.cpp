#include "StdAfx.h"
#include "Board.h"

IMPLEMENT_DYNAMIC(Board,FurniturePart)
	int Board::g_nVersion = 1;
Board::Board( void )
{
	m_dLength = 0.0;
	m_dWidth  = 0.0;
	m_dThick  = 0.0;
}

Board::~Board( void )
{

}

void Board::SetPosition( const AcGePoint3d& ptOrg )
{
	m_ptOrg = ptOrg;
}

AcGePoint3d Board::GetPosition()
{
	return m_ptOrg;
}

void Board::SetDirection( const AcGeVector3d& vecx,const AcGeVector3d& vecy,const AcGeVector3d& vecz )
{
	m_vecX = vecx;
	m_vecY = vecy;
	m_vecZ = vecz;
}

void Board::GetDirection( AcGeVector3d& vecx,AcGeVector3d& vecy,AcGeVector3d& vecz )
{
	vecx = m_vecX;
	vecy = m_vecY;
	vecz = m_vecZ;
}

void Board::SetDimension( double dLen,double dWidth,double dThick )
{
	m_dLength = dLen;
	m_dWidth  = dWidth;
	m_dThick  = dThick;
}

void Board::GetDimension( double& dLen,double& dWidth,double& dThick )
{
	dLen	= m_dLength;
	dWidth	= m_dWidth;
	dThick	= m_dThick;
}

struct resbuf* Board::PrepareXData()
{
	struct resbuf* pXdata,*pTemp;
	pXdata=FurniturePart::PrepareXData();if(pXdata==NULL)
		return NULL;

	pTemp=pXdata;
	while(pTemp->rbnext!=NULL)
		pTemp=pTemp->rbnext;
	pTemp->rbnext=acutBuildList
		(AcDb::kDxfXdAsciiString,CString(RC(Board)->m_lpszClassName),
		AcDb::kDxfXdInteger32,g_nVersion,
		AcDb::kDxfXdWorldXCoord,asDblArray(m_ptOrg),
		AcDb::kDxfXdWorldXCoord,asDblArray(m_vecX),
		AcDb::kDxfXdWorldXCoord,asDblArray(m_vecY),
		AcDb::kDxfXdWorldXCoord,asDblArray(m_vecZ),
		AcDb::kDxfXdReal,m_dLength,
		AcDb::kDxfXdReal,m_dWidth,
		AcDb::kDxfXdReal,m_dThick,
		RTNONE);

	return pXdata;
}

void Board::GiDrawFace( int iFaceNo,int color )
{
	if (iFaceNo <= 0 || color < 0 || iFaceNo > 6)
		return ;
	PrePareObject();
	FdmCoordSystem	coord = m_pFixFaceArray[iFaceNo - 1];
	coord.GrDrawCoord(3,3);
	coord.GrDrawFace(3);
}

BOOL Board::GetObjectXData( AcDbObjectId idPart,struct resbuf*& pBuf)
{
	if (!FurniturePart::GetObjectXData(idPart,pBuf))
		return FALSE;
	if (pBuf == NULL)
		return FALSE;
	pBuf = pBuf->rbnext;
	//版本号
	if (pBuf->restype == AcDb::kDxfXdInteger32)
		g_nVersion = pBuf->resval.rint;
	pBuf = pBuf->rbnext;
	//基点
	if (pBuf->restype == AcDb::kDxfXdWorldXCoord)
		m_ptOrg = asPnt3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//X向
	if (pBuf->restype == AcDb::kDxfXdWorldXCoord)
		m_vecX = asVec3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//Y向
	if (pBuf->restype == AcDb::kDxfXdWorldXCoord)
		m_vecY = asVec3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//Z向
	if (pBuf->restype == AcDb::kDxfXdWorldXCoord)
		m_vecZ = asVec3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//长度
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_dLength = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	//宽度
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_dWidth = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	//厚度
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_dThick = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	PrePareObject();
	return TRUE;
}

BOOL Board::Rebulid()
{
	if (m_id.isValid())
		UtUtil.Erase(m_id);

	AcDb3dSolid*	pSolid = new AcDb3dSolid;
	if (pSolid == NULL)
		return FALSE;
	Acad::ErrorStatus es = pSolid->createBox(m_dLength,m_dWidth,m_dThick);
	if (es != Acad::eOk)
		return FALSE;

	AcGeMatrix3d	matMove;
	AcGeVector3d	vecMove = AcGeVector3d::kXAxis*m_dLength/2+AcGeVector3d::kYAxis*m_dWidth/2+AcGeVector3d::kZAxis*m_dThick/2;
	matMove.setToTranslation(vecMove);
	pSolid->transformBy(matMove);

	//绘制板标
	double	dRadius = 200;
	double	dMinBdDim = m_dLength > m_dWidth?m_dWidth:m_dLength;
	if (dMinBdDim < dRadius)
		dRadius = dMinBdDim;
	AcDbArc*	pArc = new AcDbArc(AcGePoint3d(0,0,0),dRadius,0,PI/2);
	//创建图块
	AcDbBlockTableRecord*	newBlkRec = NULL;
	AcDbObjectId	newBlkRecId;
	DbUtil.defineNewAnonymousBlock(newBlkRec,newBlkRecId);
	AcDbObjectId	idBaseBd;
	newBlkRec->appendAcDbEntity(idBaseBd,pSolid);
	m_idBaseBd = idBaseBd;
	pSolid->close();
	AcDbObjectId	idBdMark;
	newBlkRec->appendAcDbEntity(idBdMark,pArc);
	pArc->close();
	newBlkRec->close();
	AcDbBlockReference*		pRef = new AcDbBlockReference(AcGePoint3d(0,0,0),newBlkRecId);
	AcGeMatrix3d	mat;
	mat.setToAlignCoordSys(AcGePoint3d(0,0,0),AcGeVector3d::kXAxis,AcGeVector3d::kYAxis,AcGeVector3d::kZAxis,
		m_ptOrg,m_vecX,m_vecY,m_vecZ);
	pRef->transformBy(mat);
	m_id = DbUtil.addToModelSpace(pRef);
	struct resbuf*	pXdata = PrepareXData();
	setXdata(pXdata);
	return TRUE;
}

BOOL Board::PrePareObject()
{
	if (!FurniturePart::PrePareObject())
		return FALSE;

	//构建板件六个面的坐标系
	m_pFixFaceArray.RemoveAll();
	AcGePoint3d		ptBase = m_ptOrg;
	FdmCoordSystem	coord;
	coord.SetCoordSystem(ptBase,ptBase+m_vecX*m_dLength,ptBase+m_vecZ*m_dThick,ptBase-m_vecY*100);
	m_pFixFaceArray.Add(coord);
	ptBase = m_ptOrg + m_vecX * m_dLength;
	coord.SetCoordSystem(ptBase,ptBase+m_vecY*m_dWidth,ptBase+m_vecZ*m_dThick,ptBase+m_vecX*100);
	m_pFixFaceArray.Add(coord);
	ptBase = m_ptOrg + m_vecY * m_dWidth;
	coord.SetCoordSystem(ptBase,ptBase+m_vecX*m_dLength,ptBase+m_vecZ*m_dThick,ptBase+m_vecY*100);
	m_pFixFaceArray.Add(coord);
	ptBase = m_ptOrg;
	coord.SetCoordSystem(ptBase,ptBase+m_vecY*m_dWidth,ptBase+m_vecZ*m_dThick,ptBase-m_vecX*100);
	m_pFixFaceArray.Add(coord);
	ptBase = m_ptOrg + m_vecZ * m_dThick;
	coord.SetCoordSystem(ptBase,ptBase+m_vecX*m_dLength,ptBase+m_vecY*m_dWidth,ptBase+m_vecZ*100);
	m_pFixFaceArray.Add(coord);
	ptBase = m_ptOrg;
	coord.SetCoordSystem(ptBase,ptBase+m_vecX*m_dLength,ptBase+m_vecY*m_dWidth,ptBase-m_vecZ*100);
	m_pFixFaceArray.Add(coord);

	return TRUE;
}

void Board::GetFaceNo( int& iFaceNo )
{
	int		index = 0;
	while (1)
	{
		acedCommand(RTSTR,_T("regen"),RTNONE);
		GiDrawFace(index%6+1,1);
		TCHAR	result[133] = _T("Y");
		acedInitGet(RSG_OTHER,_T("Y N"));
		int	iRet = acedGetKword(_T("\n当前面<Y>/下一个面<N>:"),result);
		if (RTCAN == iRet)
			break;
		if (RTNORM != iRet)
		{
			index ++;
			continue;
		}
		if (_tcsicmp(result,_T("Y")) == 0)
		{
			iFaceNo = index%6+1;
			break;
		}
		else 
			index ++;
	}
	acedCommand(RTSTR,_T("regen"),RTNONE);
}

void Board::GetFaceCoordSystem( int iFaceNo,AcGePoint3d& ptOrg,AcGePoint3d& ptX,AcGePoint3d& ptY,AcGePoint3d& ptZ )
{
	if (iFaceNo < 1 || iFaceNo > 6)
		return ;
	FdmCoordSystem	coord = m_pFixFaceArray[iFaceNo - 1];
	coord.GetCoordSystem(ptOrg,ptX,ptY,ptZ);
}

void Board::GetFaceDimension( int iFaceNo,double& dLen,double& dWidth )
{
	if (iFaceNo < 1 || iFaceNo > 6)
		return ;
	FdmCoordSystem	coord = m_pFixFaceArray[iFaceNo - 1];
	AcGePoint3d	ptOrg,ptX,ptY,ptZ;
	coord.GetCoordSystem(ptOrg,ptX,ptY,ptZ);
	dLen = (ptX - ptOrg).length();
	dWidth = (ptY - ptOrg).length();
}

BOOL Board::IsEmpty()
{
	return m_id.isNull();
}

BOOL Board::Clone( Board* pBoard )
{
	if (m_id.isNull())
		return FALSE;
	AcDbEntity*	pEnt = DbUtil.openAcDbEntity(m_id,AcDb::kForRead);
	AcDbEntity*	pEntClone = (AcDbEntity*)pEnt->clone();
	pEntClone->setXData(pEnt->xData());
	pEnt->close();
	AcDbObjectId	idNewBoard = DbUtil.addToModelSpace(pEntClone);
	pBoard->FromObjectId(m_id);
	pBoard->SetId(idNewBoard);
	return TRUE;
}

BOOL Board::TransformBy( AcGeMatrix3d mat )
{
	m_ptOrg.transformBy(mat);
	m_vecX.transformBy(mat);
	m_vecY.transformBy(mat);
	m_vecZ.transformBy(mat);
	for (int i = 0; i < m_pFixFaceArray.GetCount(); i ++)
	{
		FdmCoordSystem&	coord = m_pFixFaceArray[i];
		AcGePoint3d	ptOrg,ptX,ptY,ptZ;
		coord.GetCoordSystem(ptOrg,ptX,ptY,ptZ);
		ptOrg.transformBy(mat);
		ptX.transformBy(mat);
		ptY.transformBy(mat);
		ptZ.transformBy(mat);
		coord.SetCoordSystem(ptOrg,ptX,ptY,ptZ);
	}
	AcDbEntity*	pEnt = DbUtil.openAcDbEntity(m_id,AcDb::kForWrite);
	pEnt->transformBy(mat);
	pEnt->close();
	return TRUE;
}

BOOL Board::Highlight( BOOL bHighlight )
{
	AcDbEntity*	pEnt = DbUtil.openAcDbEntity(m_id,AcDb::kForWrite);
	if (pEnt == NULL)
		return FALSE;
	if (bHighlight)
		pEnt->highlight(kNullSubent,Adesk::kTrue);
	else
		pEnt->unhighlight(kNullSubent,Adesk::kTrue);
	pEnt->close();
	return TRUE;
}

BOOL Board::GetBoardExtent( AcDbExtents& extent )
{
	AcDbEntity*	pEnt = DbUtil.openAcDbEntity(m_id);
	if (pEnt == NULL)
		return FALSE;
	pEnt->getGeomExtents(extent);
	pEnt->close();
	return TRUE;
}
