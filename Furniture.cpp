#include "StdAfx.h"
#include "Furniture.h"

FdmCoordSystem::FdmCoordSystem()
{
}

FdmCoordSystem::~FdmCoordSystem()
{
}

void FdmCoordSystem::GrDrawCoord( int nArrowColor,int nSignColor ) const
{
	AcGeVector3d	vecX,vecY,vecZ;
	double			dLen,dWidth,dH;
	vecX = (m_ptX - m_ptOrg).normalize();
	dLen = (m_ptX - m_ptOrg).length();
	vecY = (m_ptY - m_ptOrg).normalize();
	dWidth = (m_ptY - m_ptOrg).length();
	vecZ = (m_ptZ - m_ptOrg).normalize();
	dH = (m_ptZ - m_ptOrg).length();
	AcGePoint3d	ptStart,ptEnd;
	//绘制X方向
	ptStart = m_ptOrg;ptEnd = m_ptOrg + vecX * (dLen + 200);
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nSignColor,0);
	ptStart = ptEnd + vecY * 20 - vecX * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nSignColor,0);
	ptStart = ptEnd - vecY * 20 - vecX * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nSignColor,0);
	ptStart = ptEnd + vecX * 200;
	ptEnd = ptStart + vecY * 85 - vecX * 50;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),2,0);
	ptEnd = ptStart + vecY * 85 + vecX * 50;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),2,0);
	ptEnd = ptStart - vecY * 85 - vecX * 50;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),2,0);
	ptEnd = ptStart - vecY * 85 + vecX * 50;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),2,0);

	//绘制Y方向
	ptStart = m_ptOrg;ptEnd = m_ptOrg + vecY * (dWidth + 200);
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nSignColor,0);
	ptStart = ptEnd + vecX * 20 - vecY * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nSignColor,0);
	ptStart = ptEnd - vecX * 20 - vecY * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nSignColor,0);
	ptStart = ptEnd + vecY * 200;
	ptEnd = ptStart + vecY * 85 - vecX * 50;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),2,0);
	ptEnd = ptStart + vecY * 85 + vecX * 50;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),2,0);
	ptEnd = ptStart - vecY * 85;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),2,0);


	//绘制正向
	ptStart = m_ptOrg + vecX * dLen /2 + vecY *dWidth/2;
	ptEnd = ptStart + vecZ * 200;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nArrowColor,0);
	ptStart = ptEnd + vecX * 20 - vecZ * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nArrowColor,0);
	ptStart = ptEnd - vecX * 20 - vecZ * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nArrowColor,0);
	ptStart = ptEnd + vecY * 20 - vecZ * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nArrowColor,0);
	ptStart = ptEnd - vecY * 20 - vecZ * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nArrowColor,0);
	ptStart = ptEnd - vecY * 20 - vecZ * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptStart + vecY * 40),nArrowColor,0);
	ptStart = ptEnd - vecX * 20 - vecZ * 100;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptStart + vecX * 40),nArrowColor,0);
}

void FdmCoordSystem::GrDrawFace( int nColor ) const
{
	AcGeVector3d	vecX,vecY,vecZ;
	double			dLen,dWidth,dH;
	vecX = (m_ptX - m_ptOrg).normalize();
	dLen = (m_ptX - m_ptOrg).length();
	vecY = (m_ptY - m_ptOrg).normalize();
	dWidth = (m_ptY - m_ptOrg).length();
	vecZ = (m_ptZ - m_ptOrg).normalize();
	dH = (m_ptZ - m_ptOrg).length();
	AcGePoint3d	ptStart,ptEnd;
	ptStart = m_ptOrg;ptEnd = m_ptOrg + vecX * dLen;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nColor,0);
	ptStart = m_ptOrg + vecX * dLen;ptEnd = m_ptOrg + vecX * dLen + vecY * dWidth;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nColor,0);
	ptStart = m_ptOrg + vecY * dWidth;ptEnd = m_ptOrg + vecX * dLen + vecY * dWidth;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nColor,0);
	ptStart = m_ptOrg;ptEnd = m_ptOrg + vecY * dWidth;
	acedGrDraw(asDblArray(ptStart),asDblArray(ptEnd),nColor,0);
}

void FdmCoordSystem::SetCoordSystem( AcGePoint3d ptOrg,AcGePoint3d ptX,AcGePoint3d ptY,AcGePoint3d ptZ )
{
	m_ptOrg = ptOrg;m_ptX = ptX;m_ptY = ptY;m_ptZ = ptZ;
}

FdmCoordSystem & FdmCoordSystem::operator=( const FdmCoordSystem &orig )
{
	m_ptOrg = orig.m_ptOrg;m_ptX = orig.m_ptX;m_ptY = orig.m_ptY;m_ptZ = orig.m_ptZ;
	return *this;
}

BOOL FdmCoordSystem::GetCoordSystem( AcGePoint3d& ptOrg,AcGePoint3d& ptX,AcGePoint3d& ptY,AcGePoint3d& ptZ )
{
	ptOrg = m_ptOrg;ptX = m_ptX;ptY = m_ptY;ptZ = m_ptZ;
	return TRUE;
}

IMPLEMENT_DYNAMIC(FurniturePart,CObject)

static CString	sFurPartApp = _T("Fur_Design");
int FurniturePart::g_nVersion = 1;
FurniturePart::FurniturePart(void)
{

}


FurniturePart::~FurniturePart(void)
{
}

struct resbuf* FurniturePart::PrepareXData()
{
	acdbRegApp(sFurPartApp);
	resbuf*	pbuf = acutBuildList
		(AcDb::kDxfRegAppName,sFurPartApp,
		AcDb::kDxfXdAsciiString,CString(RC(FurniturePart)->m_lpszClassName),
		AcDb::kDxfXdInteger32,g_nVersion,
		AcDb::kDxfXdAsciiString,m_sPartName,
		AcDb::kDxfXdAsciiString,m_sPartNo,
		RTNONE);
	return pbuf;
}

BOOL FurniturePart::setXdata( struct resbuf *xdata )
{
	if (xdata == NULL)
		return FALSE;
	AcDbEntity*		pEnt = DbUtil.openAcDbEntity(m_id,AcDb::kForWrite);
	Acad::ErrorStatus es = pEnt->setXData(xdata);
	pEnt->close();
	acutRelRb(xdata);
	return TRUE;
}

BOOL FurniturePart::GetObjectXData( AcDbObjectId idPart,struct resbuf *&pBuf )
{
	if (idPart.isNull())
		return FALSE;
	AcDbEntity*	pEnt = DbUtil.openAcDbEntity(idPart,AcDb::kForRead);
	if (pEnt == NULL)
		return FALSE;
	pBuf = pEnt->xData(sFurPartApp);
	m_pBufHead = pBuf;
	pEnt->close();
	if (pBuf == NULL)
		return FALSE;
	//类名称
	pBuf = pBuf->rbnext;
	//版本号
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdInteger32)
		g_nVersion = pBuf->resval.rint;
	pBuf = pBuf->rbnext;
	//名称
	if (pBuf->restype == AcDb::kDxfXdAsciiString)
		m_sPartName = pBuf->resval.rstring;
	pBuf = pBuf->rbnext;
	//代号
	if (pBuf->restype == AcDb::kDxfXdAsciiString)
		m_sPartNo = pBuf->resval.rstring;
	pBuf = pBuf->rbnext;
	m_id = idPart;
	return TRUE;
}

BOOL FurniturePart::fdmIsKindOf( const CRuntimeClass* pClass ) const
{
	CRuntimeClass* pRunClass = this->GetRuntimeClass();
	if (pRunClass == pClass)
		return TRUE;
	return FALSE;
}

BOOL FurniturePart::FromObjectId( AcDbObjectId idBd )
{
	struct resbuf*	pbuf = NULL;
	if (!GetObjectXData(idBd,pbuf))
		return FALSE;
	if (m_pBufHead != NULL)
	{
		acutRelRb(m_pBufHead);
		m_pBufHead = NULL;
	}
	return TRUE;
}

BOOL FurniturePart::PrePareObject()
{
	return TRUE;
}

void FurniturePart::Update()
{
	struct resbuf* pBuf = PrepareXData();
	if (pBuf == NULL)
		return ;
	AcDbEntity*	pEnt = DbUtil.openAcDbEntity(m_id,AcDb::kForWrite);
	if (pEnt == NULL)
		return;
	Acad::ErrorStatus es = pEnt->setXData(pBuf);
	es = pEnt->close();
}
