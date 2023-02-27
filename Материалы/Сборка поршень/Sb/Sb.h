
// Sb.h: главный файл заголовка для приложения PROJECT_NAME
//

#pragma once

#ifndef __AFXWIN_H__
	#error "включить pch.h до включения этого файла в PCH"
#endif

#include "resource.h"		// основные символы


// CSbApp:
// Сведения о реализации этого класса: Sb.cpp
//

class CSbApp : public CWinApp
{
public:
	CSbApp();

// Переопределение
public:
	virtual BOOL InitInstance();

// Реализация

	DECLARE_MESSAGE_MAP()
};

extern CSbApp theApp;
