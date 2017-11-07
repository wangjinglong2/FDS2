#include "StdAfx.h"
#include "CmdManager.h"
#include "Furniture.h"
#include "Hardware.h"
#include "CommonUtilFun.h"
#include "Excel/ExcelBooster.h"

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
	sPartName = _T("左侧板");
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
	sPartName = _T("右侧板");
	CommonUtilFun::CreateBoard(sPartName,sLayer,ptOrg,vecX,vecY,vecZ,dLen,dWidth,dThick);
	//创建顶板
	vecX = AcGeVector3d(1,0,0);
	vecY = AcGeVector3d(0,1,0);
	vecZ = AcGeVector3d(0,0,1);
	ptOrg = AcGePoint3d(-16,0,800);
	dLen = 1016;
	dWidth = 500;
	sPartName = _T("顶板");
	CommonUtilFun::CreateBoard(sPartName,sLayer,ptOrg,vecX,vecY,vecZ,dLen,dWidth,dThick);
	//创建底板
	ptOrg = AcGePoint3d(-16,0,-16);
	dLen = 1016;
	sPartName = _T("底板");
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
	bias.SetPartName(_T("三合一"));
	bias.SetPartNo(_T("23x34"));
	AcDbObjectIdArray	idFixBds;
	idFixBds.append(idFirstBd);
	idFixBds.append(idSecondBd);
	bias.SetFixBdArray(idFixBds);
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
}
void CCmdManager::boardReport()
{
	AcDbObjectIdArray	idPartArray;
	if (!CommonUtilFun::fds_ssGetPart(Fds::SS_FURNITUREPART,_T("请选择需要输出的板件和五金:\n"),FALSE,idPartArray))
		return;
	MSExcelBooster	excel;
	excel.set_Visible(true);
	excel.SetDisplayAlerts(false);
	if (!excel.InitExcelCOM())
		return;
	
	//excel.OpenOneWorkSheets(_T("E:\\WSoftWare_Develop\\FDS2\\template\\testtmpl.xlt"));
	excel.OpenExcelBook(_T("E:\\WJL\\githubcode\\FDS\\FDS2\\template\\test.xlt"));
	excel.SetCurWorkSheet(1);
	//excel.SetCurWorkSheet(2,_T("板件清单"));
	//excel.SetCurWorkSheet(_T("test"));
	excel.MoveTo(1,1);
	excel <<1<<2<<3<<4<<5<<endl;
	excel <<2<<4<<6<<8<<10<<endl;
	//excel.DeleteSheet(2);
	//excel.AddSheet();
	//excel.DeleteRow(1);
	//excel.DeleteRows(1,2);
	//excel.CopyRow(1,5);
	//excel <<3<<5<<7<<9<<11<<endl;
	//CString sCellPos = excel.GetCellPos(_T("测试单元格"));
	int	iRow,iColumn;
	if (excel.GetCellPos(_T("测试单元格"),iRow,iColumn))
	{
		excel.MoveTo(iRow,iColumn);
		CString sPlace1 = excel.GetCellPos(iRow,iColumn);
		CString sPlace2 = excel.GetCellPos(iRow,iColumn+1);
		excel.MergeCell(sPlace1,sPlace2);
		excel.FillMerge(sPlace1,sPlace2,_T("已替换"));
		excel.SetCellFont(sPlace1,sPlace2,RGB(255,0,0));
		excel.SetCellVerAlign(sPlace1,sPlace2,ExcelBoosterBase::VERCENALIGN);
		excel.SetCellHorAlign(sPlace1,sPlace2,ExcelBoosterBase::HORCENALIGN);
		excel.SetCellBkgColor(sPlace1,sPlace2,RGB(0,255,0));
		excel.SetCellBorder(sPlace1,sPlace2,1);
		//excel.FillMerge(_T("A5"),_T("B10"),_T("已替换"));
		//excel.InsertPicture(_T("A5"),_T("B10"),_T("C:\\Users\\admin\\Desktop\\board.jpg"),TRUE);
	}
	//开始输出物料清单
	excel.MoveTo(5,1);
	for (int i = 0; i < idPartArray.length(); i ++)
	{
		Board	board;
		board.FromObjectId(idPartArray[i]);
		double	dLen,dWidth,dH;
		board.GetDimension(dLen,dWidth,dH);
		excel << board.GetPartName()<<board.GetPartNo()<<dLen<<dWidth<<dH<<endl;
	}
	excel.SaveAsExcel(_T("E:\\WJL\\githubcode\\FDS\\FDS2\\template\\test.xlsx"));
	//excel.Quit();
}

void CCmdManager::openHardware()
{
	AcDbObjectIdArray	idPartArray;
	if (!CommonUtilFun::fds_ssGetPart(Fds::SS_HARDWARE,_T("请选择五金:\n"),FALSE,idPartArray))
		return;
	MSExcelBooster	excel;
	excel.set_Visible(true);
	excel.SetDisplayAlerts(false);
	if (!excel.InitExcelCOM())
		return;

	CString sSystemPath;
	if (!CommonUtilFun::fds_GetSystemPath(sSystemPath))
		return ;
	CString sFilePath;
	sFilePath = sSystemPath + _T("template\\test.xlt");
	if (!excel.OpenExcelBook(sFilePath))
		return;
	//excel.OpenExcelBook(_T("E:\\WJL\\githubcode\\FDS\\FDS2\\template\\test.xlt"));
	excel.SetCurWorkSheet(1);
	//excel.SetCurWorkSheet(2,_T("板件清单"));
	//excel.SetCurWorkSheet(_T("test"));
	excel.MoveTo(1,1);
	excel <<1<<2<<3<<4<<5<<endl;
	excel <<2<<4<<6<<8<<10<<endl;

	for (int i = 0; i < idPartArray.length(); i ++)
	{
		AcDbObjectId	idHw = idPartArray[i];
		FurniturePart*	pPart = NULL;
		if (!FDS_OpenPart(pPart,idHw))
			return;
		if (pPart->fdmIsKindOf(RC(HardWare)))
		{
			HardWare*	pHw = (HardWare*)pPart;
			excel << pHw->GetPartNo() << pHw->GetPartName()<<endl;
		}
		delete pPart;
		pPart = NULL;
	}
	//excel.SaveAsExcel(_T("E:\\WJL\\githubcode\\FDS\\FDS2\\template\\test.xlsx"));
	CString sSavePath;
	sSavePath.Format(_T("%s%s"),sSystemPath,_T("template\\test.xlsx"));
	excel.SaveAsExcel(sSavePath);

	return;
}
