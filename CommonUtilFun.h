#pragma once
#include "board.h"
#include "SArxUtils.h"

class CommonUtilFun
{
public:
	CommonUtilFun(void);
	~CommonUtilFun(void);

	static AcDbObjectId CreateBoard(CString sBdName,CString sLayer,AcGePoint3d ptOrg,AcGeVector3d vecX,AcGeVector3d vecY,AcGeVector3d vecZ,
		double dLen,double dWidth,double dHeight);
	static AcDbObjectId SelectObject(struct resbuf *rb,CString szPrompt,AcGePoint3d& rePt);
	static AcDbObjectId FurnitureSelect(Fds::SS_GetType ssType,CString szPrompt,AcGePoint3d& rePt );
	static void HighlightObj(AcDbObjectId idObj,BOOL bHighlight);
	static BOOL GetTwoBoardFixPos(Board& firstBd,Board& secondBd,Board*& pMainBd,Board*& pSubMainBd,int& iFaceNo,double& dCanFixLen);
	static BOOL GetCurrentView(AcDbViewTableRecord &view);
	static AcGePoint3d MidPoint(AcGePoint3d pt1,AcGePoint3d pt2);
	static BOOL TranformById(AcDbObjectId idEnt,AcGeMatrix3d mat);
	static AcDbMText * MakeMText(const TCHAR *str,const TCHAR *style, const AcGePoint3d &ptInsert, const double textHeight, const double width,const double rotate,const int color, AcDbMText::AttachmentPoint mode,const TCHAR *layer);
	static AcDbObjectId CreateTextStyle(const CString &style,const CString &fontname);
	static BOOL AddViewPortToPS( const AcGeVector3d &vecVport, const AcGePoint3d &centerPoint, const double height, const double width, AcDbObjectId &idVp, const TCHAR *layer);
	static BOOL AddViewPortToPS(const AcDbExtents &zoomExtend, const AcGeVector3d &vecVport, const AcGePoint3d &centerPoint, const double height, const double width, AcDbObjectId &idVp, const TCHAR *layer);
	static void HideWorkViewport();
	static AcDbObjectId InsertBlockToPs(AcDbObjectId blockId,const AcGePoint3d& insertPoint, const TCHAR *Layer);
	static AcDbObjectId CreateViewport(const AcGePoint3d& Position, double Width, double Height, int SetActive, const TCHAR* Layer);
	static void AgainRestoreVportOriginZoom(const ads_name Set,BOOL isZoom);
	static BOOL GetEntityMaxMinPoint(const AcDbObjectId &objId,AcGePoint3d &maxPoint,AcGePoint3d &minPoint);
	static BOOL fds_ssGetPart(Fds::SS_GetType ssType,CString szPrompt,AcDbObjectIdArray& idPartArray);
};

