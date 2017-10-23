#include "UDIDispath.h"
#include "UDWorkbooks.h"
#include "UDWorkbook.h"
#pragma once
 
/******************************************************************************
描述: 对Excel常用操作封装类，适用于Microexcel、WPSexcel表格文件
1,打开、新建、保存、另存excel文件
2,工作表的增加、删除、隐藏操作、拷贝、移动
3,单元格文本写入以及读取、各种清空
4,字体设置，前景色、背景色设置
3,行高度、列宽度设置，行列隐藏、删除
5,图片插入，包括位置以及大小
6,获取特定列的最后一个有效行索引***  个人感觉这点比较好

  背景: 基于VAB宏  C++接口
  备注: 未检测内存泄漏。 适用前确保CoInitialize(NULL)被调用。适用后释放CoUninitialize();
	
	  环境: WinXP+VC6
	  修改: QiuJL		EMAIL:282881515@163.COM		QQ:282881515
	  版本: 2014-9-30	V1.0
	  发布: CSDN 
******************************************************************************/


namespace UDExcel{
class CUDExcelApp:public CUDIDispath
{
public:
	CUDExcelApp();
	virtual ~CUDExcelApp();
	BOOL         InitServer();  
	BOOL         StopServer();
	//是否显示
	void         SetVisable(BOOL bVisable);

	//是否隐藏掉提示信息
	void         DisplayAlerts(BOOL bDisplay=FALSE);

	CUDWorkbooks GetWorkbooks();
	CUDWorkbook  GetActiveWorkbook();

    CUDSheet     GetActiveSheet();
	
public: 
}; 
}