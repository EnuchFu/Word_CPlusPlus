#include "stdafx.h"
#include "My_Word.h"
#include <io.h>

CMyWord::CMyWord()
{
	m_Wordapp = NULL;
	m_Worddocs = NULL;
	m_Worddoc = NULL;
	m_Wordrange = NULL;
	m_Wordselect = NULL;
}

CMyWord::~CMyWord()
{
	if (m_Worddoc != NULL)
	{
		m_Worddoc.Close(
			COleVariant((short)false),	// SaveChanges.
			COleVariant((short)true),	// OriginalFormat.
			COleVariant((short)false)	// RouteDocument.
			);
		m_Wordapp.Quit(
			COleVariant((short)false),	// SaveChanges.
			COleVariant((short)true),	// OriginalFormat.
			COleVariant((short)false)	// RouteDocument.
			);
		//释放内存申请资源
		m_Wordrange.ReleaseDispatch();
		m_Wordselect.ReleaseDispatch();
		m_Worddoc.ReleaseDispatch();
		m_Worddocs.ReleaseDispatch();
		m_Wordapp.ReleaseDispatch();
		m_Wordrange = NULL;
		m_Wordselect = NULL;
		m_Worddoc = NULL;
		m_Worddocs = NULL;
		m_Wordapp = NULL;
	}
}

//新建Word
bool CMyWord::CreateWord(bool isVisible)
{
	if (InitWord(isVisible))
	{
		m_Worddocs = m_Wordapp.get_Documents();
		if (!m_Worddocs.m_lpDispatch)
		{
			AfxMessageBox("Documents创建失败!", MB_OK | MB_ICONWARNING);
			return false;
		}
		COleVariant varTrue(short(1), VT_BOOL), vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		CComVariant Template(_T(""));    //没有使用WORD的文档模板
		CComVariant NewTemplate(false), DocumentType(0), Visible;

		//添加空白页
		m_Worddocs.Add(&Template, &NewTemplate, &DocumentType, &Visible);

		// 得到document变量
		m_Worddoc = m_Wordapp.get_ActiveDocument();
		if (!m_Worddoc.m_lpDispatch)
		{
			AfxMessageBox("Document获取失败!", MB_OK | MB_ICONWARNING);
			return false;
		}
		//得到selection变量
		m_Wordselect = m_Wordapp.get_Selection();
		if (!m_Wordselect.m_lpDispatch)
		{
			AfxMessageBox("Select获取失败!", MB_OK | MB_ICONWARNING);
			return false;
		}
		//得到Range变量
		m_Wordrange = m_Worddoc.Range(vOptional, vOptional);
		if (!m_Wordrange.m_lpDispatch)
		{
			AfxMessageBox("Range获取失败!", MB_OK | MB_ICONWARNING);
			return false;
		}
		return true;
	}
	return false;
}

//打开Word
bool CMyWord::OpenWord(string inWordPath, bool isVisible)
{
	//判断文件是否存在
	if (_access(inWordPath.c_str(), 0) == -1)
	{
		m_Wordapp.Quit(
			COleVariant((short)false),    // SaveChanges.
			COleVariant((short)true),            // OriginalFormat.
			COleVariant((short)false)            // RouteDocument.
			);
		m_Wordapp = NULL;
		MessageBoxA(NULL, "指定打开的文件不存在!", "提示", MB_ICONEXCLAMATION);
		return FALSE;
	}

	//判断文件是否有写入权限
	if (_access(inWordPath.c_str(), 2) == -1)
	{
		m_Wordapp.Quit(
			COleVariant((short)false),    // SaveChanges.
			COleVariant((short)true),            // OriginalFormat.
			COleVariant((short)false)            // RouteDocument.
			);
		m_Wordapp = NULL;
		MessageBoxA(NULL, "指定打开的文件没有写入权限!", "提示", MB_ICONEXCLAMATION);
		return false;
	}

	//打开Word
	if (InitWord(isVisible))
	{
		m_Worddocs = m_Wordapp.get_Documents();
		if (!m_Worddocs.m_lpDispatch)
		{
			AfxMessageBox("Documents创建失败!", MB_OK | MB_ICONWARNING);
			return false;
		}
		COleVariant vTrue((short)TRUE),
			vFalse((short)FALSE),
			vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			vZ((short)0);

		//得到document变量
		m_Worddoc = m_Worddocs.Open(
				(COleVariant)inWordPath.c_str(),        // FileName
				vTrue,            // Confirm Conversion.
				vFalse,           // ReadOnly.
				vFalse,           // AddToRecentFiles.
				vOptional,        // PasswordDocument.
				vOptional,        // PasswordTemplate.
				vOptional,        // Revert.
				vOptional,        // WritePasswordDocument.
				vOptional,        // WritePasswordTemplate.
				vOptional,        // Format. // Last argument for Word 97
				vOptional,        // Encoding // New for Word 2000/2002
				vOptional,        // Visible
				/*如下4个是word2003需要的参数。本版本是word2000。*/
				vOptional,    // OpenAndRepair
				vZ,            // DocumentDirection wdDocumentDirection LeftToRight
				vOptional,    // NoEncodingDialog
				vOptional
			);

		if (!m_Worddoc.m_lpDispatch)
		{
			AfxMessageBox("Document获取失败!", MB_OK | MB_ICONWARNING);
			return FALSE;
		}
		//得到selection变量
		m_Wordselect = m_Wordapp.get_Selection();
		if (!m_Wordselect.m_lpDispatch)
		{
			AfxMessageBox("Select获取失败!", MB_OK | MB_ICONWARNING);
			return false;
		}
		//得到Range变量
		m_Wordrange = m_Worddoc.Range(vOptional, vOptional);
		if (!m_Wordrange.m_lpDispatch)
		{
			AfxMessageBox("Range获取失败!", MB_OK | MB_ICONWARNING);
			return false;
		}
		this->m_filename = inWordPath;
		return true;
	}
	return false;
}

//保存Word
bool CMyWord::SaveWord()
{
	m_Worddoc.Save();
	return TRUE;
}

//另存Word-----覆盖替换
bool CMyWord::SaveAs(string inSavePath)
{
	try
	{
		if (strcmp(inSavePath.c_str(), this->m_filename.c_str()) != 0)
		{
			COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
			m_Worddoc.SaveAs(COleVariant(inSavePath.c_str()), covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, 
				covOptional,  covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional);
			this->m_filename = inSavePath;
			return true;
		}
		else
		{
			MessageBoxA(NULL, "输入路径和原路径相同，请直接使用SaveWord()!", "提示", MB_ICONEXCLAMATION);
			return false;
		}
	}
	catch (exception &e)
	{
		MessageBoxA(NULL, e.what(), "错误", MB_ICONEXCLAMATION);
		return false;
	}
}

//关闭Word
bool CMyWord::CloseWord()
{
	m_Worddoc.Close(
		COleVariant((short)false),	// SaveChanges.
		COleVariant((short)true),	// OriginalFormat.
		COleVariant((short)false)	// RouteDocument.
	);
	m_Wordapp.Quit(
		COleVariant((short)false),	// SaveChanges.
		COleVariant((short)true),	// OriginalFormat.
		COleVariant((short)false)	// RouteDocument.
	);

	//释放内存申请资源
	m_Wordrange.ReleaseDispatch();
	m_Wordselect.ReleaseDispatch();
	m_Worddoc.ReleaseDispatch();
	m_Worddocs.ReleaseDispatch();
	m_Wordapp.ReleaseDispatch();
	m_Wordrange = NULL;
	m_Wordselect = NULL;
	m_Worddoc = NULL;
	m_Worddocs = NULL;
	m_Wordapp = NULL;

	return true;
}

//在当前光标处插入文字
void CMyWord::WriteText(string inString)
{
	m_Wordselect.TypeText(inString.c_str());
}

//换行写字
void CMyWord::EnterLineWriteText(string inString, int nLineCount)
{
	if (nLineCount <= 0)
	{
		nLineCount = 0;
	}
	for (int i = 0; i < nLineCount; i++)
	{
		m_Wordselect.TypeParagraph();
	}
	WriteText(inString.c_str());
}

//文档结尾处写字
void CMyWord::WriteInLastLine(string inString)
{
	m_Wordrange.InsertAfter(inString.c_str());
}

//光标处插入表格
void CMyWord::InsertTable(int inRow, int inColumn, CTable0 &outTable)
{
	VARIANT vtDefault;
	COleVariant vtAuto;
	vtDefault.vt = VT_INT;
	vtDefault.intVal = 1;
	vtAuto.vt = VT_INT;
	vtAuto.intVal = 0;

	CTables0 wordtables = m_Worddoc.get_Tables();
	outTable = wordtables.Add(m_Wordselect.get_Range(), inRow, inColumn, &vtDefault, &vtAuto);
	wordtables.ReleaseDispatch();
}

//表格中写入数据
void CMyWord::WriteCellFromTable(CTable0 inTable, int inRow, int inColumn, string inString)
{
	CCell cell = inTable.Cell(inRow, inColumn);
	cell.Select(); //将光标移动到单元格
	m_Wordselect.TypeText(inString.c_str());
	cell.ReleaseDispatch();
}

//光标处插入图片（在Table中插入图片有BUG，要等~两秒）
void CMyWord::InsertPicture(string inPicturePath, bool isDelete)
{
	//判断文件是否存在
	if (_access(inPicturePath.c_str(), 0) == -1)
	{
		m_Wordapp.Quit(
			COleVariant((short)false),			// SaveChanges.
			COleVariant((short)true),            // OriginalFormat.
			COleVariant((short)false)            // RouteDocument.
			);
		m_Wordapp = NULL;
		MessageBoxA(NULL, "指定需要插入的图片不存在!", "提示", MB_ICONEXCLAMATION);
		return;
	}
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CnlineShapes cnlineShapes = m_Wordselect.get_InlineShapes();
	CnlineShape cnlineShape = cnlineShapes.AddPicture(inPicturePath.c_str(), vFalse, vTrue, vOptional);
	if (isDelete)
	{
		::DeleteFile(inPicturePath.c_str());
	}
	
}

//光标向下N行---Enter
void CMyWord::MoveSelect_E(int inN /*= 1*/)
{
	if (inN <= 0)
	{
		inN = 0;
	}
	for (int i = 0; i < inN; i++)
	{
		m_Wordselect.TypeParagraph();
	}
}

//根据书签的设置光标位置
void CMyWord::SetSelectToBookMark(string inMarkName)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	m_Wordselect.GoTo(covOptional, covOptional, covOptional, (COleVariant)inMarkName.c_str());
}

//移动光标---上下左右
void CMyWord::MoveSelect_U(int inLineCount)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	m_Wordselect.MoveUp(covOptional, COleVariant((short)inLineCount), covOptional);
}
void CMyWord::MoveSelect_D(int inLineCount)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	m_Wordselect.MoveDown(covOptional, COleVariant((short)inLineCount), covOptional);
}
void CMyWord::MoveSelect_L(int inCharCount)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	m_Wordselect.MoveLeft(covOptional, COleVariant((short)inCharCount), covOptional);
}
void CMyWord::MoveSelect_R(int inCharCount)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	m_Wordselect.MoveRight(covOptional, COleVariant((short)inCharCount), covOptional);
}

//初始化Word
bool CMyWord::InitWord(bool isVisible)
{
	if (!m_Wordapp.CreateDispatch(_T("word.application"), NULL))
	{
		AfxMessageBox("启动Word服务器失败!");
		return false;
	}
	else
	{
		m_Wordapp.put_Visible(isVisible);
		return true;
	}
}
