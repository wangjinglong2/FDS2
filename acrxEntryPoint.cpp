#include "StdAfx.h"
#include "CmdManager.h"
#include "Furniture.h"

void InitApplication();
void UnInitApplication();
void RegisterCommand();

extern "C" AcRx::AppRetCode 
	acrxEntryPoint(AcRx::AppMsgCode msg, void* pkt)
{
	switch (msg) {
	case AcRx::kInitAppMsg:
		acrxDynamicLinker->unlockApplication(pkt);
		acrxDynamicLinker->registerAppMDIAware(pkt);
		InitApplication();
		break;
	case AcRx::kUnloadAppMsg:
		UnInitApplication();
		break;
	case AcRx::kLoadDwgMsg:
		break;
	case AcRx::kUnloadDwgMsg:
		break;
	case AcRx::kInvkSubrMsg:
		break;
	}
	return AcRx::kRetOK;
}
void RegisterCommand()
{
	AcEdCommandStack*	cmdStack = acedRegCmds;
	VERIFY(cmdStack != NULL);
	cmdStack->addCommand(_T("Design"),_T("testline"),_T("testline"),ACRX_CMD_MODAL,&CCmdManager::test);
	cmdStack->addCommand(_T("Design"),_T("testboard"),_T("testboard"),ACRX_CMD_MODAL,&CCmdManager::testBoard);
	cmdStack->addCommand(_T("Design"),_T("testHardware"),_T("testHardware"),ACRX_CMD_MODAL,&CCmdManager::testHardware);
	cmdStack->addCommand(_T("Design"),_T("createframe"),_T("createframe"),ACRX_CMD_MODAL,&CCmdManager::createFrame);
	cmdStack->addCommand(_T("Design"),_T("fixhardware"),_T("fixhardware"),ACRX_CMD_MODAL,&CCmdManager::fixHardware);
	cmdStack->addCommand(_T("Design"),_T("arxhighlight"),_T("arxhighlight"),ACRX_CMD_MODAL,&CCmdManager::arxHighlight);
	cmdStack->addCommand(_T("Design"),_T("setboarducs"),_T("setboarducs"),ACRX_CMD_MODAL,&CCmdManager::setBoardUCS);
	cmdStack->addCommand(_T("Design"),_T("boardviewport"),_T("boardviewport"),ACRX_CMD_MODAL,&CCmdManager::boardViewPort);
	cmdStack->addCommand(_T("Design"),_T("boardreport"),_T("boardreport"),ACRX_CMD_MODAL,&CCmdManager::boardReport);
	cmdStack->addCommand(_T("Design"),_T("openhardware"),_T("openhardware"),ACRX_CMD_MODAL,&CCmdManager::openHardware);
}
void unRegisterCommand()
{
	acedRegCmds->removeGroup(_T("Design"));
}

void InitApplication()
{
	RegisterCommand();
	//×¢²áÎå½ğÀà

}

void UnInitApplication()
{
	unRegisterCommand();
}
