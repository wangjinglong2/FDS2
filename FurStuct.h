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
	AcGePoint3d pt[2];			//�׵���ʼ����յ�
	double dRadius,dDepth;		//�׵İ뾶�����

	CString m_sHWPartType;		//�������������
	CString m_sHWClassName;		//������������

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