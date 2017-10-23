#include "UDIDispath.h"
#include "UDWorkbooks.h"
#include "UDWorkbook.h"
#pragma once
 
/******************************************************************************
����: ��Excel���ò�����װ�࣬������Microexcel��WPSexcel����ļ�
1,�򿪡��½������桢���excel�ļ�
2,����������ӡ�ɾ�������ز������������ƶ�
3,��Ԫ���ı�д���Լ���ȡ���������
4,�������ã�ǰ��ɫ������ɫ����
3,�и߶ȡ��п�����ã��������ء�ɾ��
5,ͼƬ���룬����λ���Լ���С
6,��ȡ�ض��е����һ����Ч������***  ���˸о����ȽϺ�

  ����: ����VAB��  C++�ӿ�
  ��ע: δ����ڴ�й©�� ����ǰȷ��CoInitialize(NULL)�����á����ú��ͷ�CoUninitialize();
	
	  ����: WinXP+VC6
	  �޸�: QiuJL		EMAIL:282881515@163.COM		QQ:282881515
	  �汾: 2014-9-30	V1.0
	  ����: CSDN 
******************************************************************************/


namespace UDExcel{
class CUDExcelApp:public CUDIDispath
{
public:
	CUDExcelApp();
	virtual ~CUDExcelApp();
	BOOL         InitServer();  
	BOOL         StopServer();
	//�Ƿ���ʾ
	void         SetVisable(BOOL bVisable);

	//�Ƿ����ص���ʾ��Ϣ
	void         DisplayAlerts(BOOL bDisplay=FALSE);

	CUDWorkbooks GetWorkbooks();
	CUDWorkbook  GetActiveWorkbook();

    CUDSheet     GetActiveSheet();
	
public: 
}; 
}