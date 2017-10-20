// FurnitureDesign.cpp : 定义 DLL 的初始化例程。
//

#include "stdafx.h"
#include "FurnitureDesign.h"
#include "CmdManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

BEGIN_MESSAGE_MAP(CFurnitureDesignApp, CWinApp)
END_MESSAGE_MAP()

CFurnitureDesignApp::CFurnitureDesignApp()
{

}
CFurnitureDesignApp theApp;

BOOL CFurnitureDesignApp::InitInstance()
{
	CWinApp::InitInstance();

	if (!AfxSocketInit())
	{
		AfxMessageBox(IDP_SOCKETS_INIT_FAILED);
		return FALSE;
	}

	return TRUE;
}


