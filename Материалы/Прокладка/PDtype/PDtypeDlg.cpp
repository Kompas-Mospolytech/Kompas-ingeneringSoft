//***************************************************************************************************
//Программа построения прокладки в компас 3Д
//Автор: Куклин Данил Александрович 31.12.2002 г.Москва
// Дата last изменения: 14.11.2022
//***************************************************************************************************




// PDtypeDlg.cpp: файл реализации
//

#include "pch.h"
#include "framework.h"
#include "PDtype.h"
#include "PDtypeDlg.h"
#include "afxdialogex.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\kAPI5.tlb"
#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\ksConstants3D.tlb"
#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\ksConstants.tlb"

using namespace Kompas6API5;

// Диалоговое окно CAboutDlg используется для описания сведений о приложении

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Данные диалогового окна
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // поддержка DDX/DDV

// Реализация
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// Диалоговое окно CPDtypeDlg



CPDtypeDlg::CPDtypeDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_PDTYPE_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CPDtypeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_BUTTON1, m_start);
}

BEGIN_MESSAGE_MAP(CPDtypeDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CPDtypeDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// Обработчики сообщений CPDtypeDlg

BOOL CPDtypeDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Добавление пункта "О программе..." в системное меню.

	// IDM_ABOUTBOX должен быть в пределах системной команды.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Задает значок для этого диалогового окна.  Среда делает это автоматически,
	//  если главное окно приложения не является диалоговым
	SetIcon(m_hIcon, TRUE);			// Крупный значок
	SetIcon(m_hIcon, FALSE);		// Мелкий значок

	// TODO: добавьте дополнительную инициализацию

	return TRUE;  // возврат значения TRUE, если фокус не передан элементу управления
}

void CPDtypeDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// При добавлении кнопки свертывания в диалоговое окно нужно воспользоваться приведенным ниже кодом,
//  чтобы нарисовать значок.  Для приложений MFC, использующих модель документов или представлений,
//  это автоматически выполняется рабочей областью.

void CPDtypeDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // контекст устройства для рисования

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Выравнивание значка по центру клиентского прямоугольника
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Нарисуйте значок
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// Система вызывает эту функцию для получения отображения курсора при перемещении
//  свернутого окна.
HCURSOR CPDtypeDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CPDtypeDlg::OnBnClickedButton1()
{

	CoInitialize(NULL);
	Kompas6API5::KompasObjectPtr kompass;
	HRESULT hRes;
	hRes = kompass.GetActiveObject(L"Kompas.Application.5");
	if (FAILED(hRes))
		kompass.CreateInstance(L"Kompas.Application.5");
	kompass->Visible = true;

	Kompas6API5::ksDocument3DPtr idoc3D = kompass->Document3D();
	idoc3D->Create(false, true);
	idoc3D = kompass->ActiveDocument3D();


	//эскиз1 
	Kompas6API5::ksPartPtr part;
	part = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	Kompas6API5::ksEntityPtr ISketchEntity = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	Kompas6API5::ksSketchDefinitionPtr sketchDef = ISketchEntity->GetDefinition();
	sketchDef->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity->Create();
	Kompas6API5::ksDocument2DPtr doc2d = sketchDef->BeginEdit();

	Kompas6API5::ksMathematic2DPtr Mathematic2D;
	Mathematic2D = *new Kompas6API5::ksMathematic2DPtr(kompass->GetMathematic2D());

	//построение сопряг. дуги
	doc2d->ksArcByPoint(40, 40,12.5, 40, 52.5, 52.5, 40, -1, 1);
	doc2d->ksArcByPoint(-40, 40, 12.5, -52.5, 40, -40, 52.5, -1, 1);
	doc2d->ksArcByPoint(-40,-40, 12.5, -52.5, -40, -40, -52.5, 1, 1);
	doc2d->ksArcByPoint(40, -40, 12.5, 40, -52.5, 52.5, -40, 1, 1);

	//сопрягаемые отрезки
	doc2d->ksLineSeg(-40, 52.5, 40, 52.5, 1);
	doc2d->ksLineSeg(-52.5, 40, -52.5, -40, 1);
	doc2d->ksLineSeg(-40, -52.5, 40, -52.5, 1);
	doc2d->ksLineSeg(52.5, -40, 52.5, 40, 1);
	
	
	sketchDef->EndEdit();

	//выдавливание
	Kompas6API5::ksEntityPtr BossExtr = part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion); //получение интерфейса объекта 
	Kompas6API5::ksBossExtrusionDefinitionPtr BossExtrDef = BossExtr->GetDefinition(); //получаем интерфейс параметров этой операции 
	BossExtrDef->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 3, 0, FALSE); //установление параметров 
	BossExtrDef->directionType* (Kompas6Constants3D::dtNormal); 
	BossExtrDef->SetSketch(ISketchEntity); //чо као именно делат дац дац 
	BossExtr->Create();

	//эскиз2 отверстия BIG
	Kompas6API5::ksPartPtr part2;
	part2 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	Kompas6API5::ksEntityPtr ISketchEntity2 = part2->NewEntity(Kompas6Constants3D::o3d_sketch);
	Kompas6API5::ksSketchDefinitionPtr sketchDef2 = ISketchEntity2->GetDefinition();
	sketchDef2->SetPlane(part2->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity2->Create();
	Kompas6API5::ksDocument2DPtr doc2d2 = sketchDef2->BeginEdit();

	Kompas6API5::ksMathematic2DPtr Mathematic2D2;
	Mathematic2D2 = *new Kompas6API5::ksMathematic2DPtr(kompass->GetMathematic2D());

	doc2d2->ksCircle(0, 0, 40, 1);

	sketchDef2->EndEdit();

	//вырез выдавливание отверстие BIG
	Kompas6API5::ksEntityPtr CutExtr = part2->NewEntity(Kompas6Constants3D::o3d_cutExtrusion); //получение интерфейса объекта 
	Kompas6API5::ksCutExtrusionDefinitionPtr CutExtrDef = CutExtr -> GetDefinition(); //получаем интерфейс параметров этой операции 
	CutExtrDef->SetSideParam(TRUE, Kompas6Constants3D::etThroughAll, 3, 0, FALSE); 
	CutExtrDef->directionType* (Kompas6Constants3D::dtNormal); 
	CutExtrDef->SetSketch(ISketchEntity2); 
	CutExtr->Create(); 

	//эскиз3 отверстия MINI
	Kompas6API5::ksPartPtr part3;
	part3 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	Kompas6API5::ksEntityPtr ISketchEntity3 = part3->NewEntity(Kompas6Constants3D::o3d_sketch);
	Kompas6API5::ksSketchDefinitionPtr sketchDef3 = ISketchEntity3->GetDefinition();
	sketchDef3->SetPlane(part3->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity3->Create();
	Kompas6API5::ksDocument2DPtr doc2d3 = sketchDef3->BeginEdit();

	Kompas6API5::ksMathematic2DPtr Mathematic2D3;
	Mathematic2D3 = *new Kompas6API5::ksMathematic2DPtr(kompass->GetMathematic2D());

	doc2d3->ksCircle(40, 40, 5, 1);
	doc2d3->ksCircle(-40, 40, 5, 1);
	doc2d3->ksCircle(-40, -40, 5, 1);
	doc2d3->ksCircle(40, -40, 5, 1);

	sketchDef3->EndEdit();

	//вырез выдавливанием отверстий MINI
	Kompas6API5::ksEntityPtr CutExtr2 = part3->NewEntity(Kompas6Constants3D::o3d_cutExtrusion); //получение интерфейса объекта 
	Kompas6API5::ksCutExtrusionDefinitionPtr CutExtrDef2 = CutExtr2->GetDefinition(); //получаем интерфейс параметров этой операции 
	CutExtrDef2->SetSideParam(TRUE, Kompas6Constants3D::etThroughAll, 3, 0, FALSE);
	CutExtrDef2->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef2->SetSketch(ISketchEntity3);
	CutExtr2->Create();

	kompass.Detach();
}
