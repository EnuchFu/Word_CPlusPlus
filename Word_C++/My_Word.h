
#ifndef My_Word_H_INCLUDED
#define My_Word_H_INCLUDED

#pragma once

#include "CApplication.h"
#include "CDocument0.h"
#include "CDocuments.h"
#include "CnlineShape.h"
#include "CnlineShapes.h"
#include "CRange.h"
#include "CSelection.h"
#include "CShape.h"
#include "CShapes.h"
#include "CTable0.h"
#include "CTables0.h"
#include "CTextFrame.h"
#include "CCell.h"
#include "CCells.h"
#include "CBookmark0.h"
#include "CBookmarks.h"

#include <string>
#include <vector>

#define DllExport __declspec(dllexport)

using namespace std;

class DllExport CMyWord
{
public:
	CMyWord();
	virtual ~CMyWord();

public:

	//新建Word
	bool CreateWord(bool isVisible);

	//打开Word
	bool OpenWord(string inWordPath, bool isVisible = false);

	//保存Word
	bool SaveWord();

	//另存Word-----覆盖替换
	bool SaveAs(string inSavePath);

	//关闭Word
	bool CloseWord();

	//当前光标处写文本
	void WriteText(string inString);

	//换N行写字
	void EnterLineWriteText(string inString, int nLineCount = 1);

	//文档结尾处写字
	void WriteInLastLine(string inString);

	//光标处插入表格
	void InsertTable(int inRow, int inColumn, CTable0 &outTable);

	//表格中写入数据
	void WriteCellFromTable(CTable0 inTable, int inRow, int inColumn, string inString);

	//光标处插入图片
	void InsertPicture(string inPicturePath, bool isDelete = true);

	//光标向下N行
	void MoveSelect_E(int inN = 1);

	//根据书签的设置光标位置
	void SetSelectToBookMark(string inMarkName);

	//移动光标---上下左右
	void MoveSelect_U(int inLineCount);
	void MoveSelect_D(int inLineCount);
	void MoveSelect_L(int inCharCount);
	void MoveSelect_R(int inCharCount);

private:
	CApplication	m_Wordapp;
	CDocuments		m_Worddocs;
	CDocument0		m_Worddoc;
	CRange			m_Wordrange;
	CSelection		m_Wordselect;

	string m_filename;

	//初始化Word
	bool InitWord(bool isVisible);

};

#endif