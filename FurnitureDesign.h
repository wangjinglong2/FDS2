// FurnitureDesign.h : FurnitureDesign DLL ����ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CFurnitureDesignApp
// �йش���ʵ�ֵ���Ϣ������� FurnitureDesign.cpp
//

class CFurnitureDesignApp : public CWinApp
{
public:
	CFurnitureDesignApp();

// ��д
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};
