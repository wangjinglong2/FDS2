#pragma once

class FdmCoordSystem
{
public:
	FdmCoordSystem();
	virtual ~FdmCoordSystem();
	FdmCoordSystem &operator=(const FdmCoordSystem &orig);
	void SetCoordSystem(AcGePoint3d ptOrg,AcGePoint3d ptX,AcGePoint3d ptY,AcGePoint3d ptZ);
	BOOL GetCoordSystem(AcGePoint3d& ptOrg,AcGePoint3d& ptX,AcGePoint3d& ptY,AcGePoint3d& ptZ);
	void GrDrawCoord(int nArrowColor,int nSignColor) const;
	void GrDrawFace(int nColor) const;
private:
	AcGePoint3d m_ptOrg;
	AcGePoint3d m_ptX;
	AcGePoint3d m_ptY;
	AcGePoint3d m_ptZ;
};

class FurniturePart:public CObject
{
	DECLARE_DYNAMIC(FurniturePart)
	static int g_nVersion;//°æ±¾ºÅ
public:
	FurniturePart(void);
	virtual ~FurniturePart(void);
	BOOL FromObjectId( AcDbObjectId idBd );
	BOOL fdmIsKindOf(const CRuntimeClass* pClass) const;
	virtual BOOL Rebulid()=0;
	void Update();

	AcDbObjectId GetId(){return m_id;}
	void SetPartName(CString sVal){m_sPartName = sVal;};
	CString GetPartName(){return m_sPartName;};
	void SetPartNo(CString sVal){m_sPartNo = sVal;};
	CString GetPartNo(){return m_sPartNo;};


protected:
	virtual struct resbuf* PrepareXData();
	virtual BOOL setXdata(struct resbuf *xdata);
	virtual BOOL GetObjectXData(AcDbObjectId idPart,struct resbuf *&xdata);
	virtual BOOL PrePareObject();
	void SetId(AcDbObjectId idTemp){m_id = idTemp;};
protected:
	AcDbObjectId	m_id;
	CString			m_sPartName;
	CString			m_sPartNo;
	struct resbuf*	m_pBufHead;
};