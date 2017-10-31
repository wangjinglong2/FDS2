#include "StdAfx.h"
#include "Hardware.h"
#include "CommonUtilFun.h"

int	HardWare::g_nVersion = 1;
IMPLEMENT_DYNAMIC(HardWare,FurniturePart)
	HardWare::HardWare()
{
	m_bCallDwgBlock = FALSE;
	m_sPartClsName = CString(RC(HardWare)->m_lpszClassName);
}

HardWare::~HardWare()
{

}

struct resbuf* HardWare::PrepareXData()
{
	struct resbuf* pXdata,*pTemp;
	pXdata=FurniturePart::PrepareXData();if(pXdata==NULL)
		return NULL;

	pTemp=pXdata;
	while(pTemp->rbnext!=NULL)
		pTemp=pTemp->rbnext;
	pTemp->rbnext = acutBuildList
		(AcDb::kDxfXdAsciiString,CString(RC(HardWare)->m_lpszClassName),
		AcDb::kDxfXdInteger32,g_nVersion,
		AcDb::kDxfXdWorldXCoord,asDblArray(m_ptBase),
		AcDb::kDxfXdWorldXCoord,asDblArray(m_ptX),
		AcDb::kDxfXdWorldXCoord,asDblArray(m_ptY),
		AcDb::kDxfXdWorldXCoord,asDblArray(m_ptZ),
		AcDb::kDxfXdInteger16,m_bCallDwgBlock,
		AcDb::kDxfXdAsciiString,m_sDwgName,
		RTNONE);
	//连接板件的句柄
	while (pTemp->rbnext != NULL)
		pTemp = pTemp->rbnext;
	for (int i = 0; i < m_idFixBdArray.length(); i ++)
	{
		AcDbObjectId	idBd = m_idFixBdArray[i];;
		AcDbHandle	handle = idBd.handle();
		CString		sHandle;
		handle.getIntoAsciiBuffer(sHandle.GetBuffer(_MAX_DIR));
		pTemp->rbnext = acutBuildList(AcDb::kDxfXdAsciiString,sHandle,RTNONE);
		pTemp = pTemp->rbnext;
	}

	return pXdata;
}

void HardWare::SetPosition( AcGePoint3d ptOrg,AcGePoint3d ptX,AcGePoint3d ptY,AcGePoint3d ptZ )
{
	m_ptBase = ptOrg;m_ptX = ptX;m_ptY = ptY;m_ptZ = ptZ;
}

BOOL HardWare::GetObjectXData( AcDbObjectId idPart,struct resbuf *&pBuf )
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
		m_ptBase = asPnt3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//X点
	if (pBuf->restype == AcDb::kDxfXdWorldXCoord)
		m_ptX = asPnt3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//Y点
	if (pBuf->restype == AcDb::kDxfXdWorldXCoord)
		m_ptY = asPnt3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//Z点
	if (pBuf->restype == AcDb::kDxfXdWorldXCoord)
		m_ptZ = asPnt3d(pBuf->resval.rpoint);
	pBuf = pBuf->rbnext;
	//是否从dwg插入
	if (pBuf->restype == AcDb::kDxfXdInteger16)
		m_bCallDwgBlock = pBuf->resval.rint;
	pBuf = pBuf->rbnext;
	//dwg文件名
	if (pBuf->restype == AcDb::kDxfXdAsciiString)
		m_sDwgName = pBuf->resval.rstring;
	pBuf = pBuf->rbnext;

	PrePareObject();
	return TRUE;
}

BOOL HardWare::PrePareObject()
{
	if (!FurniturePart::PrePareObject())
		return FALSE;

	return TRUE;
}

int	BiasConnecter::g_nVersion = 1;
IMPLEMENT_SERIAL(BiasConnecter,HardWare,1)
BiasConnecter::BiasConnecter(void)
{
	m_sDwgName = _T("Biasconnecter.dwg");
	m_biasData.m_d1 = 15;
	m_biasData.m_depth1 = 13.5;
	m_biasData.m_d2 = 8;
	m_biasData.m_depth2 = 20;
	m_biasData.m_d3 = 10;
	m_biasData.m_depth3 = 12;
	m_biasData.m_depth4 = 25;
	m_biasData.m_offset = 9;
	m_sPartClsName = CString(RC(BiasConnecter)->m_lpszClassName);
}


BiasConnecter::~BiasConnecter(void)
{
}

BOOL BiasConnecter::Rebulid()
{
	CString	sDwgPath;
	getDwgPath(sDwgPath);
	AcDbObjectId	idBlkRecd;
	if (Acad::eOk != addBlkRecordFromDWG(sDwgPath,idBlkRecd))
		return FALSE;

	AcDbBlockReference*	pBlockRef = new AcDbBlockReference(AcGePoint3d(0,0,0),idBlkRecd);
	pBlockRef->setColorIndex(3);
	pBlockRef->transformBy(m_matrix);
	m_id = DbUtil.addToModelSpace(pBlockRef);
	struct resbuf*	pXdata = PrepareXData();
	setXdata(pXdata);
	return TRUE;
}

BOOL BiasConnecter::getDwgPath(CString& sDwgPath)
{
	CString	sAppPath = _T("E:\\WSoftWare_Develop\\FDS\\hwdwg\\");
	sDwgPath = sAppPath + m_sDwgName;
	return TRUE;
}
Acad::ErrorStatus  BiasConnecter::addBlkRecordFromDWG(CString szDwgPathName,AcDbObjectId& idBlockRecord,CString szBlkName)
{
	Acad::ErrorStatus es=Acad::eOk;

	if (szBlkName.IsEmpty())
	{
		szBlkName = szDwgPathName;
		szBlkName.Replace(_T(".dwg"),_T(""));
		int nFounde = szBlkName.ReverseFind('\\');
		szBlkName = szBlkName.Right(szBlkName.GetLength()-nFounde-1);
	}

	if (!HasBlock(szBlkName,idBlockRecord))
	{
		AcDbDatabase *pDb=new AcDbDatabase(Adesk::kFalse);
		if (pDb==NULL)
		{
			return Acad::eOutOfMemory;
		}
		es = pDb->readDwgFile(szDwgPathName);
		if(	es!=Acad::eOk)
		{
			delete pDb;
			acutPrintf(_T("读取五金图块文件%s失败!\n"),szDwgPathName);
			return es;
		}
		es = acdbCurDwg()->insert(idBlockRecord,szBlkName,pDb);
		if(	es!=Acad::eOk)
		{
			delete pDb;
			acutPrintf(_T("插入图块%s失败!\n"),szBlkName);
			return es;
		}
		delete pDb;
	}

	return es;
}
BOOL BiasConnecter::HasBlock(CString szRecordName,AcDbObjectId & id)
{
	AcDbBlockTable* blkTbl=NULL;
	Acad::ErrorStatus es;
	AcDbDatabase* pDb = acdbCurDwg();
	es = pDb->getBlockTable(blkTbl, AcDb::kForRead);
	if (es != Acad::eOk) 
	{
		return FALSE;
	}
	BOOL bHas = blkTbl->has(szRecordName);
	if (bHas)
	{
		bHas = blkTbl->getAt(szRecordName,id);
	}
	blkTbl->close();
	return bHas;
}

BOOL BiasConnecter::PrepareHwData(Board& firstBd,Board& secondBd,BOOL bPosFace,double dLayDist)
{
	if (firstBd.IsEmpty()|| secondBd.IsEmpty()|| dLayDist < 0)
		return FALSE;
	//根据两个板件的位置关系获取三合一的安装位置
	Board	*pMainBd = NULL,*pSubMainBd = NULL;
	int		iFaceNo = -1;
	double	dCanFixLen = 0.0;
	if (!CommonUtilFun::GetTwoBoardFixPos(firstBd,secondBd,pMainBd,pSubMainBd,iFaceNo,dCanFixLen)&&!
		CommonUtilFun::GetTwoBoardFixPos(secondBd,firstBd,pMainBd,pSubMainBd,iFaceNo,dCanFixLen))
		return FALSE;

	if (pMainBd == NULL || pSubMainBd == NULL || iFaceNo <=0 || dCanFixLen <= 0.0)
		return FALSE;
	AcGePoint3d	ptOrg,ptX,ptY,ptZ;
	pMainBd->GetFaceCoordSystem(iFaceNo,ptOrg,ptX,ptY,ptZ);
	AcGeVector3d	vecX,vecY,vecZ;
	vecX = (ptX - ptOrg).normalize();
	vecY = (ptY - ptOrg).normalize();
	vecZ = (ptZ - ptOrg).normalize();
	AcGePoint3d	ptBiasBase;
	ptBiasBase = ptOrg + vecX * dLayDist + vecY * m_biasData.m_offset - vecZ*m_biasData.m_depth4;
	AcGeMatrix3d	mat;
	mat.setToAlignCoordSys(AcGePoint3d(0,0,0),AcGeVector3d::kXAxis,AcGeVector3d::kYAxis,AcGeVector3d::kZAxis,
		ptBiasBase,vecX,vecZ,-vecY);
	m_matrix = mat;
	m_ptBase = ptBiasBase;
	m_ptX = m_ptBase + vecX * 1;
	m_ptY = m_ptBase + vecZ * 1;
	m_ptZ = m_ptBase + (-vecY) * 1;

	return TRUE;
}

struct resbuf* BiasConnecter::PrepareXData()
{
	struct resbuf* pXdata,*pTemp;
	pXdata=HardWare::PrepareXData();if(pXdata==NULL)
		return NULL;

	pTemp=pXdata;
	while(pTemp->rbnext!=NULL)
		pTemp=pTemp->rbnext;
	pTemp->rbnext = acutBuildList
		(AcDb::kDxfXdAsciiString,CString(RC(BiasConnecter)->m_lpszClassName),
		AcDb::kDxfXdInteger32,g_nVersion,
		AcDb::kDxfXdReal,m_biasData.m_d1,
		AcDb::kDxfXdReal,m_biasData.m_depth1,
		AcDb::kDxfXdReal,m_biasData.m_d2,
		AcDb::kDxfXdReal,m_biasData.m_depth2,
		AcDb::kDxfXdReal,m_biasData.m_d3,
		AcDb::kDxfXdReal,m_biasData.m_depth3,
		AcDb::kDxfXdReal,m_biasData.m_depth4,
		AcDb::kDxfXdReal,m_biasData.m_offset,
		RTNONE);

	return pXdata;
}

BOOL BiasConnecter::GetObjectXData( AcDbObjectId idPart,struct resbuf *&pBuf )
{
	if (!HardWare::GetObjectXData(idPart,pBuf))
		return FALSE;
	if (pBuf == NULL)
		return FALSE;
	pBuf = pBuf->rbnext;
	//版本号
	if (pBuf->restype == AcDb::kDxfXdInteger32)
		g_nVersion = pBuf->resval.rint;
	pBuf = pBuf->rbnext;
	//偏心件尺寸
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_d1 = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_depth1 = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_d2 = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_depth2 = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_d3 = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_depth3 = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_depth4 = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	if (pBuf->restype == AcDb::kDxfXdReal)
		m_biasData.m_offset = pBuf->resval.rreal;
	pBuf = pBuf->rbnext;
	PrePareObject();

	return TRUE;
}

BOOL BiasConnecter::PrePareObject()
{
	if (!HardWare::PrePareObject())
		return FALSE;
	return TRUE;
}
