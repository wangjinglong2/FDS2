#include "StdAfx.h"
#include "CommonUtilFun.h"


CommonUtilFun::CommonUtilFun(void)
{
}


CommonUtilFun::~CommonUtilFun(void)
{
}

AcDbObjectId CommonUtilFun::CreateBoard( CString sBdName,CString sLayer,AcGePoint3d ptOrg,AcGeVector3d vecX,
	AcGeVector3d vecY,AcGeVector3d vecZ, double dLen,double dWidth,double dHeight )
{
	Board	board;
	board.SetPosition(ptOrg);
	board.SetDimension(dLen,dWidth,dHeight);
	board.SetDirection(vecX,vecY,vecZ);
	board.Rebulid();
	return board.GetId();
}

AcDbObjectId CommonUtilFun::SelectObject( struct resbuf *rb,CString szPrompt,AcGePoint3d& rePt )
{
	struct resbuf *rbGetObject=NULL;
	ads_name ent;
	AcDbObjectId Id;
	ads_point pt;
	while(1)
	{
		ads_initget(0,_T("E"));
		int re=acedEntSel(szPrompt,ent,pt);
		if(re==RTCAN||re==RTKWORD)
			return NULL;
		else if(re!=RTNORM)
			continue;
		if (rb==NULL)
		{
			break;
		}
		rbGetObject=acdbEntGet(ent);
		BOOL bWillBreak=FALSE;
		if(rbGetObject&&rbGetObject->rbnext&&rbGetObject->rbnext->restype==0)
		{
			struct resbuf *rbTemp=rb;
			while(rbTemp)
			{
				if(rbTemp->restype==0)
				{
					if(!_tcscmp(rbGetObject->rbnext->resval.rstring,rbTemp->resval.rstring))
					{
						rePt=asPnt3d(pt);
						bWillBreak=TRUE;
						break;
					}
				}
				rbTemp=rbTemp->rbnext;
			}
		}
		if(bWillBreak)
			break;
		else
		{
			ads_ssfree(ent);
			ads_printf(_T("\n选择不正确！"));
		}
	}
	Acad::ErrorStatus es = acdbGetObjectId(Id,ent);
	if(es != Acad::eOk)
	{
		ads_ssfree(ent);
		return NULL;
	}
	ads_ssfree(ent);
	return Id;
}

AcDbObjectId CommonUtilFun::FurnitureSelect(Fds::SS_GetType ssType,CString szPrompt,AcGePoint3d& rePt )
{
	CString	sObjName;
	if (ssType == Fds::SS_FURNITUREPART)
		sObjName = _T("FurniturePart");
	else if (ssType == Fds::SS_BOARD)
		sObjName = _T("Board");
	else if (ssType == Fds::SS_HARDWARE)
		sObjName = _T("HardWare");

	struct resbuf *rbGetObject=NULL;
	ads_name ent;
	AcDbObjectId Id;
	ads_point pt;
	while(1)
	{
		ads_initget(0,_T("E"));
		int re=acedEntSel(szPrompt,ent,pt);
		if(re==RTCAN||re==RTKWORD)
			return NULL;
		else if(re!=RTNORM)
			continue;
		struct resbuf *apps=NULL; 
		apps=acutBuildList(RTSTR,_T("Fur_Design"),RTNONE);
		rbGetObject=acdbEntGetX(ent,apps);
		acutRelRb(apps);
		struct resbuf *rbTemp = rbGetObject;
		BOOL bWillBreak=FALSE;
		while (rbTemp)
		{
			if(rbTemp->restype==1000)
			{
				if(_tcscmp(rbTemp->resval.rstring,sObjName)==0)
				{
					rePt=asPnt3d(pt);
					bWillBreak=TRUE;
					break;
				}
			}
			rbTemp=rbTemp->rbnext;
		}
		if(bWillBreak)
			break;
		else
		{
			ads_ssfree(ent);
			ads_printf(_T("\n选择不正确！"));
		}
	}
	Acad::ErrorStatus es = acdbGetObjectId(Id,ent);
	if(es != Acad::eOk)
	{
		ads_ssfree(ent);
		return NULL;
	}
	ads_ssfree(ent);
	return Id;
}

void CommonUtilFun::HighlightObj( AcDbObjectId idObj,BOOL bHighlight )
{
	AcDbEntity*	pEnt = DbUtil.openAcDbEntity(idObj,AcDb::kForWrite);
	if (pEnt == NULL)
		return;
	if (bHighlight)
		pEnt->highlight(kNullSubent,Adesk::kTrue);
	else
		pEnt->unhighlight(kNullSubent,Adesk::kTrue);
	pEnt->close();
}

//获取两个板件的安装位置
BOOL CommonUtilFun::GetTwoBoardFixPos( Board& firstBd,Board& secondBd,Board*& pMainBd,Board*& pSubMainBd,int& iFaceNo,double& dCanFixLen )
{
	if (firstBd.IsEmpty() || secondBd.IsEmpty())
		return FALSE;
	iFaceNo = 0;
	dCanFixLen = 0.0;
	//第二个板件的第五面坐标
	AcGePoint3d	ptOrg2,ptX2,ptY2,ptZ2;
	secondBd.GetFaceCoordSystem(5,ptOrg2,ptX2,ptY2,ptZ2);
	//第五面的平面
	AcGePlane	plane(ptX2,ptOrg2,ptY2);
	//第五面的平面
	AcGePoint3d	ptOrg3,ptX3,ptY3,ptZ3;
	secondBd.GetFaceCoordSystem(6,ptOrg3,ptX3,ptY3,ptZ3);
	AcGePlane	plane2(ptX3,ptOrg3,ptY3);

	//第一个板件边面各点投影至第五面
	AcGePoint3d	ptOrg,ptX,ptY,ptZ;
	for (int i = 1; i <= 4; i ++)
	{
		firstBd.GetFaceCoordSystem(i,ptOrg,ptX,ptY,ptZ);
		AcGePoint3d	ptBase = ptOrg;
		ptOrg = ptOrg.orthoProject(plane);
		ptX = ptX.orthoProject(plane);
		ptY = ptY.orthoProject(plane);

		AcDbExtents	extent1;
		extent1.addPoint(ptOrg);
		extent1.addPoint(ptX);
		extent1.addPoint(ptY);

		AcDbExtents	extent2;
		extent2.addPoint(ptOrg2);
		extent2.addPoint(ptX2);
		extent2.addPoint(ptY2);

		//获取相交区域的长度和宽度
		AcGePoint3d	ptMin1,ptMax1,ptMin2,ptMax2,ptMax,ptMin;
		ptMin1 = extent1.minPoint();
		ptMax1 = extent1.maxPoint();
		ptMin2 = extent2.minPoint();
		ptMax2 = extent2.maxPoint();

		ptMax.x = ptMax1.x < ptMax2.x?ptMax1.x:ptMax2.x;
		ptMax.y = ptMax1.y < ptMax2.y?ptMax1.y:ptMax2.y;
		ptMax.z = ptMax1.z < ptMax2.z?ptMax1.z:ptMax2.z;

		ptMin.x = ptMin1.x > ptMin2.x?ptMin1.x:ptMin2.x;
		ptMin.y = ptMin1.y > ptMin2.y?ptMin1.y:ptMin2.y;
		ptMin.z = ptMin1.z > ptMin2.z?ptMin1.z:ptMin2.z;

		AcGeVector3d	vecextent = (ptMax - ptMin);
		AcGeVector3d	vecX = (ptX - ptOrg).normalize();
		AcGeVector3d	vecY = (ptY - ptOrg).normalize();
		double	dConcideLen,dConcideWidth;
		dConcideLen = fabs(vecextent.dotProduct(vecX));
		dConcideWidth = fabs(vecextent.dotProduct(vecY));
		dConcideLen = dConcideLen > dConcideWidth?dConcideLen:dConcideWidth;
		dConcideWidth = dConcideLen > dConcideWidth?dConcideWidth:dConcideLen;
		double	dFaceLen,dFaceWidth;
		firstBd.GetFaceDimension(i,dFaceLen,dFaceWidth);
		if (!fds_equal(dFaceWidth,dConcideWidth))
			continue;
		//判断面到第二块板件的距离
		AcGePoint3d	ptTemp = ptOrg.orthoProject(plane);
		AcGePoint3d	ptTemp2 = ptOrg.orthoProject(plane2);
		if (fds_equal(ptBase.distanceTo(ptTemp),0.0) || fds_equal(ptBase.distanceTo(ptTemp2),0.0))
		{
			iFaceNo = i;
			dCanFixLen = dConcideLen;
			pMainBd = &firstBd;
			pSubMainBd = &secondBd;
			return TRUE;
		}
	}
	return FALSE;
}
BOOL CommonUtilFun::GetCurrentView(AcDbViewTableRecord &view)
{
	struct resbuf var;
	struct resbuf WCS, UCS, DCS;
	WCS.restype = RTSHORT;
	WCS.resval.rint = 0;
	UCS.restype = RTSHORT;
	UCS.resval.rint = 1;
	DCS.restype = RTSHORT;
	DCS.resval.rint = 2;

	acedGetVar(_T("VIEWMODE"), &var);//当前视口的显示模式　0－关闭　1-激活透视视图　
	//2-打开前剪切　4-打开后剪切　8-UCS模式　16-剪切前不在视线内
	view.setPerspectiveEnabled(var.resval.rint & 1);
	view.setFrontClipEnabled(var.resval.rint & 2 ? true : false);
	view.setBackClipEnabled(var.resval.rint & 4 ? true : false);
	view.setFrontClipAtEye(!(var.resval.rint & 16));

	acedGetVar(_T("BACKZ"), &var);//后裁剪的平面位移量
	view.setBackClipDistance(var.resval.rreal);

	acedGetVar(_T("VIEWCTR"), &var);//视图中心
	ads_trans(var.resval.rpoint, &UCS, &DCS, NULL, var.resval.rpoint);
	view.setCenterPoint(AcGePoint2d(var.resval.rpoint[X],
		var.resval.rpoint[Y]));

	acedGetVar(_T("FRONTZ"), &var);//当前视口的前置裁剪平面偏移量
	view.setFrontClipDistance(var.resval.rreal);

	acedGetVar(_T("LENSLENGTH"), &var);//当前视口透视使用的镜头的焦距长度（毫米）
	view.setLensLength(var.resval.rreal);

	acedGetVar(_T("TARGET"), &var);//当前视口目标点的位置
	ads_trans(var.resval.rpoint, &UCS, &WCS, NULL, var.resval.rpoint);
	view.setTarget(AcGePoint3d(var.resval.rpoint[X], var.resval.rpoint[Y],
		var.resval.rpoint[Z]));

	acedGetVar(_T("VIEWDIR"), &var);//当前视口的视线方向
	ads_trans(var.resval.rpoint, &UCS, &WCS, TRUE, var.resval.rpoint);
	view.setViewDirection(AcGeVector3d(var.resval.rpoint[X],
		var.resval.rpoint[Y], var.resval.rpoint[Z]));

	acedGetVar(_T("VIEWSIZE"), &var);//视图高度	
	view.setHeight(var.resval.rreal);

	acedGetVar(_T("VIEWTWIST"), &var);	
	view.setViewTwist(var.resval.rreal);
	CRect rect;
	acedGetAcadDwgView()->GetClientRect(&rect);
	if(rect.Height() == 0) return FALSE;

	double width =((double)rect.Width()/(double)rect.Height())*view.height();
	view.setWidth(width);

	return TRUE;
}

AcGePoint3d CommonUtilFun::MidPoint( AcGePoint3d pt1,AcGePoint3d pt2 )
{
	AcGePoint3d	ptMid;
	ptMid.x = (pt1.x + pt2.x)/2;
	ptMid.y = (pt1.y + pt2.y)/2;
	ptMid.z = (pt1.z + pt2.z)/2;
	return ptMid;
}

BOOL CommonUtilFun::TranformById( AcDbObjectId idEnt,AcGeMatrix3d mat)
{
	if (idEnt.isNull())
		return FALSE;
	AcDbEntity*	pEnt = NULL;
	pEnt = DbUtil.openAcDbEntity(idEnt,AcDb::kForWrite);
	if (pEnt == NULL)
		return FALSE;
	pEnt->transformBy(mat);
	pEnt->close();
	return TRUE;
}

AcDbMText * CommonUtilFun::MakeMText(const TCHAR *str,const TCHAR *style,
	const AcGePoint3d &ptInsert, const double textHeight,
	const double width,const double rotate,const int color,
	AcDbMText::AttachmentPoint mode,const TCHAR *layer)
{
	AcDbObjectId idStyle=AcDbObjectId::kNull;
	idStyle=CreateTextStyle(_T("FDS_TEXT"),_T("新宋体"));

	AcDbMText *pMText=new AcDbMText;
	pMText->setRotation(rotate);
	pMText->setTextHeight(textHeight);
	pMText->setContents(str);
	if(idStyle!=AcDbObjectId::kNull)
		pMText->setTextStyle(idStyle);
	pMText->setWidth(width);
	pMText->setLocation(ptInsert);
	if(color<256 && color>0)
		pMText->setColorIndex(color);
	pMText->setLayer(layer);
	pMText->setAttachment(mode);

	return pMText;
}
AcDbObjectId CommonUtilFun::CreateTextStyle(const CString &style,const CString &fontname)
{
	AcDbTextStyleTable  *StyleTable;
	AcDbTextStyleTableRecord *StyleTableRecord;
	AcDbObjectId StyleId;

	//设定字型
	acdbCurDwg()->getTextStyleTable(StyleTable,AcDb::kForRead);
	if(StyleTable->has(style)==Adesk::kTrue)
	{
		StyleTable->getAt(style,StyleTableRecord,AcDb::kForWrite);
		StyleId=StyleTableRecord->objectId();
		StyleTableRecord->close();
	}
	else
	{
		StyleTableRecord = new AcDbTextStyleTableRecord();
		StyleTableRecord->setName(style);
		CString lfontname=fontname;
		lfontname.MakeLower();
		if(lfontname.Find(_T(".shx"))>=0)
		{
			StyleTableRecord->setFileName(fontname);
		}
		else{
			if(_tcsicmp(fontname,_T("黑体"))==0)
			{
				StyleTableRecord->setFont(fontname, 0, 0, 134, 2);
			}
			else
				StyleTableRecord->setFont(fontname, 0, 0, 136, 2);
		}
		StyleTable->upgradeOpen();
		StyleTable->add(StyleId,StyleTableRecord);
		StyleTable->downgradeOpen();
		StyleTableRecord->close();
	}
	StyleTable->close();

	return StyleId;
}
BOOL CommonUtilFun::AddViewPortToPS(
	const AcGeVector3d &vecVport,
	const AcGePoint3d &centerPoint,
	const double height,
	const double width,
	AcDbObjectId &idVp,
	const TCHAR *layer) 
{
	AcDbViewport* pVp = new AcDbViewport; 
	pVp->setLayer(layer);
	idVp=DbUtil.addToPaperSpace(pVp,false);
	if(idVp!=AcDbObjectId::kNull)
	{
		pVp->setCenterPoint(centerPoint);
		pVp->setHeight(height);
		pVp->setWidth(width);
		pVp->setViewDirection(vecVport); 
		pVp->setOn(); 
		pVp->close(); 
		return TRUE;
	}
	else
	{
		delete pVp;
		return FALSE;
	}

}
BOOL CommonUtilFun::AddViewPortToPS(const AcDbExtents &zoomExtend,
	const AcGeVector3d &vecVport,
	const AcGePoint3d &centerPoint,
	const double height,
	const double width,
	AcDbObjectId &idVp,
	const TCHAR *layer) 
{

	AcDbViewport* pVp = new AcDbViewport; 
	pVp->setLayer(layer);
	idVp=DbUtil.addToPaperSpace(pVp,false);


	if(idVp!=AcDbObjectId::kNull)//idVp.isValid()==true
	{
		pVp->setCenterPoint(centerPoint);
		pVp->setHeight(height);
		pVp->setWidth(width);
		//Set View direction
		pVp->setViewDirection(vecVport); 


		//Assume target point is WCS (0,0,0)
		AcGeMatrix3d ucsMatrix;
		AcGePlane XYPlane(AcGePoint3d(0,0,0), vecVport);
		ucsMatrix.setToWorldToPlane(XYPlane);

		// Get a rectangular window in viewport_T('s x-y plane that represents drawing')s extents
		AcGePoint2d maxExt, minExt;
		AcDbExtents ext=zoomExtend;
		ext.transformBy(ucsMatrix);
		AcGePoint3d max = ext.maxPoint();
		AcGePoint3d min = ext.minPoint();
		maxExt[X] = max[X];
		maxExt[Y] = max[Y];
		minExt[X] = min[X];
		minExt[Y] = min[Y];

		pVp->setOff();

		AcGePoint2d ViewCenter;
		ViewCenter[X] = (maxExt[X]+minExt[X])/2;
		ViewCenter[Y] = (maxExt[Y]+minExt[Y])/2;
		double height,vscale,zscale;
		double vheight=pVp->height();
		double vwidth=pVp->width();
		vscale=vwidth/vheight;
		zscale=(maxExt[X]-minExt[X])/(maxExt[Y]-minExt[Y]);

		height=vscale>zscale?(maxExt[Y]-minExt[Y]):((maxExt[X]-minExt[X])*vheight/vwidth);
		pVp->setViewCenter(ViewCenter);
		pVp->setViewHeight(height+0.5);
		pVp->setOn(); 
		pVp->close(); 
		return TRUE;
	}
	else
	{
		delete pVp;
		return FALSE;
	}
}
//隐藏cad工作视口
void CommonUtilFun::HideWorkViewport()
{   
	AcDbBlockTable *pBlockTable;
	acdbHostApplicationServices()->workingDatabase() ->getSymbolTable(pBlockTable, AcDb::kForRead);
	AcDbBlockTableRecord *pBlockTableRecord;
	pBlockTable->getAt(ACDB_PAPER_SPACE, pBlockTableRecord, AcDb::kForRead);
	pBlockTable->close();
	AcDbBlockTableRecordIterator *pblockiter;
	CString sLayer;
	if (Acad::eOk == pBlockTableRecord->newIterator( pblockiter))
	{   
		AcDbObjectId objid;
		for ( pblockiter->start(); !pblockiter->done();pblockiter->step()) 
		{
			if (Acad::eOk == pblockiter->getEntityId( objid))		
			{    
				AcDbEntity *pEnt=NULL;
				if(acdbOpenAcDbEntity(pEnt,objid,AcDb::kForWrite)==Acad::eOk)
				{
					if(pEnt->isA()==AcDbViewport::desc())
					{
						struct  resbuf  *pRb;
						pRb=pEnt->xData(_T("VIEWSTATE"));
						if(pRb!=NULL)
						{
							AcDbViewport *pVp=(AcDbViewport *)pEnt;
							pVp->setOff();
							sLayer= pVp->layer();
							//pVp->freezeLayersInViewport();
							//pVp->thawAllLayersInViewport();
							pVp->setVisibility(AcDb::kInvisible);
						}
					}
					pEnt->close();
				}

			}
		}
		delete pblockiter;
	}
	pBlockTableRecord->close();
	if (sLayer.IsEmpty())
		return ;
	AcDbLayerTable *pLayerTable;
	acdbHostApplicationServices()->workingDatabase()
		->getSymbolTable(pLayerTable, AcDb::kForWrite);   
	AcDbLayerTableRecord *pLayerTableRecord=NULL;
	Acad::ErrorStatus es=pLayerTable->getAt(sLayer,pLayerTableRecord,AcDb::kForWrite);
	pLayerTableRecord->setIsOff(true);
	pLayerTable->close();
	pLayerTableRecord->close();
}

AcDbObjectId CommonUtilFun::InsertBlockToPs(AcDbObjectId blockId,const AcGePoint3d& insertPoint, const TCHAR *Layer)
{
	AcDbObjectId insertId;
	AcDbBlockReference *pBlkRef=new AcDbBlockReference();
	pBlkRef->setBlockTableRecord(blockId);
	pBlkRef->setPosition(insertPoint);	
	pBlkRef->setLayer(Layer);

	AcDbBlockTable *pBlockTable;
	acdbCurDwg()->getBlockTable(pBlockTable,AcDb::kForRead);

	AcDbBlockTableRecord *pRecord;
	pBlockTable->getAt(ACDB_PAPER_SPACE,pRecord,AcDb::kForWrite);
	pBlockTable->close();

	pRecord->appendAcDbEntity(insertId,pBlkRef);
	pBlkRef->close();
	pRecord->close();

	return insertId;
}

AcDbObjectId CommonUtilFun::CreateViewport(const AcGePoint3d& Position, double Width, double Height, int SetActive, const TCHAR* Layer)
{
	//CSwitchWorkingDatabase switchdb(g_pDb);
	AcDbBlockTable *BlockTable;
	AcDbBlockTableRecord *BlockRecord;
	AcDbViewport  *Viewport;
	AcDbObjectId  objId;

	acdbCurDwg()->getBlockTable(BlockTable,AcDb::kForRead);
	BlockTable->getAt(ACDB_PAPER_SPACE, BlockRecord, AcDb::kForWrite);
	BlockTable->close();

	Viewport = new AcDbViewport();

	Viewport->setCenterPoint(Position);
	Viewport->setHeight(Height);
	Viewport->setWidth(Width);
	Viewport->setViewDirection(AcGeVector3d(0,0,1));
	//fdm_SetObjLayer(Viewport,Layer);
	Viewport->setUcsIconVisible();		//add this row

	BlockRecord->appendAcDbEntity(objId, Viewport);
	BlockRecord->close();

	//if(g_pDb==NULL)
	//{
	//	if(SetActive) Viewport->setOn();
	//	else Viewport->setOff();
	//}

	Viewport->close();

	return objId;
}

void CommonUtilFun::AgainRestoreVportOriginZoom(const ads_name Set,BOOL isZoom)
{
	ads_name Ent,sset,ename;
	long Number,i,j,index=0,viewnum=0;
	AcDbObjectId ObjId;
	AcDbObjectIdArray ObjIdArray;
	AcGePoint3d maxpt,minpt;  
	BOOL isAdd=TRUE;
	int vpNumber = 28;
	ads_command(RTSTR,_T("PSPACE"),0);
	ads_sslength(Set, &Number);
	ads_ssadd(NULL,NULL,sset);
	for(i=0L; i<Number; i++)
	{
		if(ads_ssname(Set, i, Ent)==RTNORM)
		{
			if(acdbGetObjectId(ObjId,Ent)==Acad::eOk) ObjIdArray.append(ObjId);
		}
	}  
	if(isZoom)
	{
		CommonUtilFun::GetEntityMaxMinPoint(ObjId,maxpt,minpt);
		ads_command(RTSTR,_T("ZOOM"),RTSTR,_T("W"),RT3DPOINT,asDblArray(minpt),RT3DPOINT,asDblArray(maxpt),0);
	}
	//保证要显示的视窗是开并活动的
	ads_command(RTSTR,_T("MVIEW"),RTSTR,_T("ON"),RTPICKS,sset,RTSTR,_T(""),0);
	ads_ssfree(sset);
	ads_command(RTSTR,_T("MSPACE"),0);
	//for(i=0L; i<index; i++) //只恢复显示前60个视窗
	//{
	//	ads_ssname(Set, i, Ent);
	//	RestoreVport(Ent);
	//}  
	//SwitchToPaperSpace();
	if(isZoom) 
	{
		ads_command(RTSTR,_T("ZOOM"),RTSTR,_T("E"),0);
	}
}
BOOL CommonUtilFun::GetEntityMaxMinPoint(const AcDbObjectId &objId,AcGePoint3d &maxPoint,AcGePoint3d &minPoint)
{
	BOOL state=FALSE;
	AcDbEntity *pEnt;
	if(acdbOpenObject(pEnt,objId,AcDb::kForRead)==Acad::eOk)
	{
		if(pEnt->isKindOf(AcDbOrdinateDimension::desc())) //极坐标尺寸
		{
			double maxx1,maxy1,minx1,miny1;

			maxx1=maxy1=minx1=miny1=0;
			AcDbOrdinateDimension *pOrdinate=AcDbOrdinateDimension::cast(pEnt);
			minPoint=pOrdinate->definingPoint();
			maxPoint=pOrdinate->leaderEndPoint();
			maxx1=maxPoint.x;
			maxx1=max(maxx1,minPoint.x);
			maxy1=maxPoint.y;
			maxy1=max(maxy1,minPoint.y);
			minx1=minPoint.x;
			minx1=min(minx1,maxPoint.x);
			miny1=minPoint.y;
			miny1=min(miny1,maxPoint.y);
			minPoint.set(minx1,miny1,0);
			maxPoint.set(maxx1,maxy1,0);
		}
		else
		{
			AcDbExtents extents;
			pEnt->getGeomExtents(extents);
			maxPoint=extents.maxPoint();
			minPoint=extents.minPoint();
		}
		state=TRUE;
		pEnt->close();
	}

	return state;
}
BOOL CommonUtilFun::fds_ssGetPart(Fds::SS_GetType ssType,CString szPrompt,AcDbObjectIdArray& idPartArray)
{
	CString	sObjName;
	if (ssType == Fds::SS_FURNITUREPART)
		sObjName = _T("FurniturePart");
	else if (ssType == Fds::SS_BOARD)
		sObjName = _T("Board");
	else if (ssType == Fds::SS_HARDWARE)
		sObjName = _T("HardWare");

	struct resbuf *rbGetObject=NULL;
	AcDbObjectId Id;
	ads_name ss;
	int	iRet = acedSSGet(NULL,NULL,NULL,NULL,ss);
	if (iRet != RTNORM)
		return FALSE;
	long	nCount = 0;
	iRet = acedSSLength(ss,&nCount);
	if (iRet != RTNORM)
	{
		acedSSFree(ss);
		return FALSE;
	}
	for (long i = 0;i < nCount; i ++)
	{
		ads_name	ent;
		acedSSName(ss,i,ent);
		struct resbuf *apps=NULL; 
		apps=acutBuildList(RTSTR,_T("Fur_Design"),RTNONE);
		rbGetObject=acdbEntGetX(ent,apps);
		acutRelRb(apps);
		struct resbuf *rbTemp = rbGetObject;
		BOOL bFind = FALSE;
		while (rbTemp)
		{
			if(rbTemp->restype==1000)
			{
				if(_tcscmp(rbTemp->resval.rstring,sObjName)==0)
				{
					bFind=TRUE;
					break;
				}
			}
			rbTemp=rbTemp->rbnext;
		}
		if(!bFind)
		{
			ads_ssfree(ent);
			continue;
		}
		acdbGetObjectId(Id,ent);
		idPartArray.append(Id);
		ads_ssfree(ent);
	}
	ads_ssfree(ss);
	return TRUE;
}