#include "StdAfx.h"
#include "CmdManager.h"
#include "Furniture.h"
#include "Hardware.h"
#include "CommonUtilFun.h"

CCmdManager::CCmdManager(void)
{
}


CCmdManager::~CCmdManager(void)
{
}

void CCmdManager::test()
{
	AcDbLine*	pline = new AcDbLine(AcGePoint3d(0,0,0),AcGePoint3d(100,100,100));
	DbUtil.addToModelSpace(pline);
}
void CCmdManager::testBoard()
{
	Board	board;
	board.SetPosition(AcGePoint3d(0,0,0));
	board.SetDimension(500,500,16);
	board.SetDirection(AcGeVector3d(1,0,0),AcGeVector3d(0,1,0),AcGeVector3d(0,0,1));
	board.Rebulid();
}

void CCmdManager::testHardware()
{
	BiasConnecter	bias;
	bias.Rebulid();
}

void CCmdManager::createFrame()
{
	//创建左侧板
	CString	sPartName,sLayer;
	sPartName = _T("");
	sLayer = _T("");
	AcGePoint3d	ptOrg;
	ptOrg = AcGePoint3d(0,0,0);
	AcGeVector3d	vecX,vecY,vecZ;
	vecX = AcGeVector3d(0,1,0);
	vecY = AcGeVector3d(0,0,1);
	vecZ = AcGeVector3d(-1,0,0);
	double	dLen = 500,dWidth = 800,dThick = 16;
	CommonUtilFun::CreateBoard(sPartName,sLayer,ptOrg,vecX,vecY,vecZ,dLen,dWidth,dThick);
	//创建右侧板
	vecX = AcGeVector3d(0,1,0);
	vecY = AcGeVector3d(0,0,1);
	vecZ = AcGeVector3d(-1,0,0);
	ptOrg = AcGePoint3d(1000,0,0);
	CommonUtilFun::CreateBoard(sPartName,sLayer,ptOrg,vecX,vecY,vecZ,dLen,dWidth,dThick);
	//创建顶板
	vecX = AcGeVector3d(1,0,0);
	vecY = AcGeVector3d(0,1,0);
	vecZ = AcGeVector3d(0,0,1);
	ptOrg = AcGePoint3d(-16,0,800);
	dLen = 1016;
	dWidth = 500;
	CommonUtilFun::CreateBoard(sPartName,sLayer,ptOrg,vecX,vecY,vecZ,dLen,dWidth,dThick);
	//创建底板
	ptOrg = AcGePoint3d(-16,0,-16);
	dLen = 1016;
	CommonUtilFun::CreateBoard(sPartName,sLayer,ptOrg,vecX,vecY,vecZ,dLen,dWidth,dThick);
}

void CCmdManager::fixHardware()
{
	//选择第一块板件
	AcGePoint3d		ptSel;
	AcDbObjectId idFirstBd = CommonUtilFun::FurnitureSelect(Fds::SS_BOARD,_T("\n请选择板件:"),ptSel);
	CommonUtilFun::HighlightObj(idFirstBd,TRUE);
	AcDbObjectId idSecondBd = CommonUtilFun::FurnitureSelect(Fds::SS_BOARD,_T("\n请选择板件:"),ptSel);
	CommonUtilFun::HighlightObj(idSecondBd,TRUE);
	Board	firstBoard,secondBoard;
	firstBoard.FromObjectId(idFirstBd);
	secondBoard.FromObjectId(idSecondBd);

	BiasConnecter	bias;
	bias.PrepareHwData(firstBoard,secondBoard,FALSE,200);
	bias.Rebulid();
	CommonUtilFun::HighlightObj(idFirstBd,FALSE);
	CommonUtilFun::HighlightObj(idSecondBd,FALSE);
}

void CCmdManager::arxHighlight()
{
	ads_name ssName, ssTemp;
	acedSSAdd(NULL, NULL, ssTemp);

	acedSSGet(NULL, NULL, NULL, NULL, ssName);
	long len;
	acedSSLength(ssName, &len);

	for (int i = 0; i < len; i++)
	{
		ads_name ent;
		acedSSName(ssName, i, ent);

		AcDbObjectId objid;
		acdbGetObjectId(objid, ent);
		AcDbEntity *pent=NULL;
		acdbOpenAcDbEntity(pent, objid, AcDb::kForRead);
		if (pent->isKindOf(AcDbText::desc()))
		{
			acedSSAdd(ent, ssTemp, ssTemp);
		}
		pent->close();
	}
	acedSSFree(ssName);
	acedSSSetFirst(ssTemp, NULL); 
}

void CCmdManager::setBoardUCS()
{
	//选择第一块板件
	AcGePoint3d		ptSel;
	AcDbObjectId idFirstBd = CommonUtilFun::FurnitureSelect(Fds::SS_BOARD,_T("\n请选择板件:"),ptSel);
	CommonUtilFun::HighlightObj(idFirstBd,TRUE);
	Board	board;
	board.FromObjectId(idFirstBd);

	int	iFaceNo;
	board.GetFaceNo(iFaceNo);
	if (iFaceNo <= 0)
		return;
	AcGePoint3d	ptOrg,ptX,ptY,ptZ;
	board.GetFaceCoordSystem(iFaceNo,ptOrg,ptX,ptY,ptZ);
	double	dFaceLen,dFaceWidth;
	board.GetFaceDimension(iFaceNo,dFaceLen,dFaceWidth);
	ads_command(RTSTR,_T("UCS"),RTSTR,
		_T("3POINT"),RT3DPOINT,asDblArray(ptOrg),
		RT3DPOINT,asDblArray(ptX),
		RT3DPOINT,asDblArray(ptY),
		0);
	AcGeVector3d vecX = acdbCurDwg()->ucsxdir();
	double	dAngle = -(vecX.angleTo(AcGeVector3d::kXAxis));
	AcDbViewTableRecord view;
	CommonUtilFun::GetCurrentView(view);
	AcGePoint3d	ptMid = CommonUtilFun::MidPoint(ptX,ptY);
	view.setTarget(ptMid); 
	view.setCenterPoint(AcGePoint2d(0,0)); 
	view.setViewDirection((ptZ-ptOrg).normalize());
	view.setWidth(dFaceLen); 
	view.setHeight(dFaceWidth); 
	view.setViewTwist(dAngle);
	acedSetCurrentView( &view, NULL ); 
	TCHAR	result[133];
	acedGetString(0,_T("请输入任意字符继续:\n"),result);

	//ads_command(RTSTR,_T("ucs"),RTSTR,_T("p"),RTNONE);
	acedCommand(RTSTR,_T("pspace"),RTNONE);
}

void CCmdManager::boardViewPort()
{
	ads_command(RTSTR,_T("mspace"),RTNONE);
	ads_command(RTSTR,_T("ucs"),RTSTR,_T("w"),0);
	AcGePoint3d		ptSel;
	AcDbObjectId idFirstBd = CommonUtilFun::FurnitureSelect(Fds::SS_BOARD,_T("\n请选择板件:"),ptSel);
	CommonUtilFun::HighlightObj(idFirstBd,TRUE);
	Board	board;
	board.FromObjectId(idFirstBd);
	Board	cloneBoard;
	board.Clone(&cloneBoard);
	AcGeMatrix3d	mat;
	mat.setToTranslation(AcGeVector3d(-10000,-10000,0));
	cloneBoard.TransformBy(mat);
	AcGePoint3d	ptOrg = cloneBoard.GetPosition();

	AcDbMText	*pMText = CommonUtilFun::MakeMText(_T("技术要求："),_T("仿宋体"),ptOrg,20,300,0,2,AcDbMText::kMiddleCenter,_T("0"));
	DbUtil.addToModelSpace(pMText);

	ads_command(RTSTR,_T("PSPACE"),0);
	double	Totalx = 1200;
	double  Totaly = 1200;
	struct resbuf *rb;
	rb=ads_newrb(RTREAL);
	double	fxp;
	TCHAR	xp[20], FilePath[255];
	if(ads_getvar(_T("VIEWSIZE"),rb)==RTNORM)
	{
		fxp=rb->resval.rreal/(Totaly*2);
		if(fxp<1)
		{
			_stprintf(xp, _T("%.6fx"), fxp);
			ads_command(RTSTR,_T("ZOOM"),RTSTR,xp,0);
		}
	}
	ads_relrb(rb);
	//指定基点
	AcGePoint3d	ptBase(0,0,0);
	if(ads_findfile(_T("B4_HOR.DWG"),FilePath)!=RTNORM)
	{
		acutPrintf(_T("未找到文件%s\n"),_T("B4_HOR.DWG"));
		return;
	}
	//插入图框
	AcDbDatabase FileData(Adesk::kFalse);
	if (FileData.readDwgFile(FilePath) != Acad::eOk)
	{
		CString strerr;
		strerr.Format(_T("读取%s文件失败！"),FilePath);
		AfxMessageBox(strerr);
		return;
	}
	AcDbObjectId	Id;
	acdbCurDwg()->insert(Id, _T("*U"), &FileData, Adesk::kFalse);
	AcDbObjectId EntId = CommonUtilFun::InsertBlockToPs(Id,ptBase,_T("0"));
	//ads_name	ent;
	//acdbGetAdsName(ent, EntId);
	AcDbBlockReference   *BlkRef;
	AcDbBlockTableRecord *BlockRecord;
	AcDbBlockTableRecordIterator *pIterator;

	if(acdbOpenObject(BlkRef,EntId,AcDb::kForWrite)!=Acad::eOk) return;
	AcDbObjectId RecordId = BlkRef->blockTableRecord();
	AcGeMatrix3d Matrix = BlkRef->blockTransform();
	AcDbEntity*	 pEnt = NULL;
	AcGePoint3d pt1,pt2,Center;
	double x,y,UP_DIST;
	if(acdbOpenObject(BlockRecord,RecordId,AcDb::kForRead)==Acad::eOk)
	{
		BlockRecord->newIterator(pIterator);
		for(pIterator->start(); !pIterator->done(); pIterator->step())
		{
			pIterator->getEntity(pEnt, AcDb::kForRead);
			AcDbExtents	extents;
			if(pEnt->colorIndex()==201)					//矩形绘图区域
			{
				pEnt->getGeomExtents(extents);
				pt1 = extents.maxPoint();	pt1.transformBy(Matrix);
				pt2 = extents.minPoint();	pt2.transformBy(Matrix);
				x = fabs(pt1.x-pt2.x);
				y = fabs(pt1.y-pt2.y);
				Center.set((pt1.x+pt2.x)/2, (pt1.y+pt2.y)/2, 0);
				pEnt->upgradeOpen();
				pEnt->setVisibility(AcDb::kInvisible);
				pEnt->downgradeOpen();
			}     
			pEnt->close();
		}
		delete pIterator;
		BlockRecord->close();
	}
	BlkRef->close(); 
	UP_DIST=15.8334;

	CommonUtilFun::HideWorkViewport();
	AcDbObjectId ObjId = CommonUtilFun::CreateViewport(Center,x,y,0,_T("0"));
	AcDbViewport *vp;
	if(acdbOpenObject(vp, ObjId, AcDb::kForWrite)==Acad::eOk)
	{
		AcDbExtents extents;
		vp->setViewDirection(AcGeVector3d(0,0,1));
		AcDbHandle	arxHandle;
		vp->getAcDbHandle(arxHandle);
		vp->setVisibility(AcDb::kInvisible,Adesk::kTrue);
		vp->getGeomExtents(extents);
		vp->setOff();
		//arxHandle.getIntoAsciiBuffer(pdata->VpHandle);
		vp->close();
	}
	ads_name	ent;
	acdbGetAdsName(ent, EntId);
	CommonUtilFun::AgainRestoreVportOriginZoom(ent,TRUE);
	if(acdbOpenObject(vp, ObjId, AcDb::kForWrite)==Acad::eOk)
	{
		AcDbExtents extents;
		vp->setViewDirection(AcGeVector3d(0,0,1));
		AcDbHandle	arxHandle;
		vp->getAcDbHandle(arxHandle);
		vp->setVisibility(AcDb::kVisible,Adesk::kTrue);
		vp->getGeomExtents(extents);
		vp->setOn();
		//arxHandle.getIntoAsciiBuffer(pdata->VpHandle);
		vp->close();
	}
	ads_command(RTSTR,_T("mspace"),0);
	AcGePoint3d	maxpt,minpt;
	CommonUtilFun::GetEntityMaxMinPoint(cloneBoard.GetId(),maxpt,minpt);
	AcDbExtents	extents;
	extents.addPoint(maxpt);
	extents.addPoint(minpt);
	extents.expandBy(AcGeVector3d(200,200,200));
	extents.expandBy(AcGeVector3d(-200,-200,-200));
	maxpt = extents.maxPoint();
	minpt = extents.minPoint();
	ads_command(RTSTR,_T("ZOOM"),RTSTR,_T("W"),RT3DPOINT,asDblArray(minpt),RT3DPOINT,asDblArray(maxpt),0);
	ads_command(RTSTR,_T("PSPACE"),0);
	AcGePoint3d ptorg;
	ptorg[0]=0.0;
	ptorg[1]=0.0;
	ptorg[2]=0.0;

	AcGePoint3d ptmax;
	ptmax[0]=ptorg[0]+297;
	ptmax[1]=ptorg[1]+210+1*20;
	ptmax[2]=0.0;
	
	//AcDbObjectId layerId = AcDbObjectId::kNull;
	//AcDbLayerTable* pLayerTable;
	//if(acdbCurDwg()->getLayerTable(pLayerTable,AcDb::kForRead) == Acad::eOk)
	//{
	//	if (pLayerTable->has(_T("0")))
	//	{
	//		pLayerTable->getAt(_T("0"),layerId);
	//	}

	//	pLayerTable->close();
	//}

	//AcDbBlockTable* pBlockTbl;
	//if(acdbCurDwg()->getBlockTable(pBlockTbl,AcDb::kForRead) == Acad::eOk)
	//{
	//	AcDbBlockTableRecord* pBlockTableRec;
	//	pBlockTbl->getAt(ACDB_PAPER_SPACE,pBlockTableRec,AcDb::kForRead);
	//	pBlockTbl->close();

	//	AcDbBlockTableRecordIterator* pLtr;
	//	pBlockTableRec->newIterator(pLtr);
	//	AcDbEntity* pEnt;
	//	for (pLtr->start();!pLtr->done();pLtr->step())
	//	{
	//		pLtr->getEntity(pEnt,AcDb::kForWrite);

	//		if (pEnt->layerId() == layerId && pEnt->isKindOf(AcDbViewport::desc())) //位于图层DWGFRAMELAYER上，且是视口
	//		{
	//			pEnt->setVisibility(AcDb::kVisible);
	//		}

	//		pEnt->close();
	//	}

	//	delete pLtr;
	//	pBlockTableRec->close();
	//}
}
