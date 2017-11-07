namespace Fds
{
	enum SS_GetType
	{
		SS_FURNITUREPART,
		SS_BOARD,
		SS_HARDWARE
	};
}
typedef struct tagHole
{
	AcGePoint3d pt[2];			//孔的起始点和终点
	double dRadius,dDepth;		//孔的半径和深度

	CString m_sHWPartType;		//五金件的零件类型
	CString m_sHWClassName;		//五金件的类名称

	tagHole()
	{
		dRadius = 0.0;
		dDepth = 0.0;
	}

	tagHole &operator=(const tagHole &orig)
	{
		for (int i = 0; i < 2;i ++)
			pt[i] = orig.pt[i];
		dRadius = orig.dRadius;
		dDepth = orig.dDepth;
		m_sHWPartType = orig.m_sHWPartType;
		m_sHWClassName = orig.m_sHWClassName;
		return (*this);
	}
}Hole;