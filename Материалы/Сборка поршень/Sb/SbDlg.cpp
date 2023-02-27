
// SbDlg.cpp: файл реализации
//

#include "pch.h"
#include "framework.h"
#include "Sb.h"
#include "SbDlg.h"
#include "afxdialogex.h"
#import "E:\kompas\Bin\kAPI5.tlb"
#include "E:\kompas\SDK\Include\ksConstants3D.h"
#include "E:\kompas\SDK\Include\ksConstants.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace Kompas6API5;
KompasObjectPtr kompass;
ksPartPtr Part, Partsb;
ksDocument3DPtr iDoc3D, iDoc3Dsb;

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


// Диалоговое окно CSbDlg



CSbDlg::CSbDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SB_DIALOG, pParent)
    , dl_shtock(210)
    , dl_cylinder(155)
    , h_ver_kriska(15)
    , h_nedokriska(25)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CSbDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Text(pDX, IDC_EDIT1, dl_shtock);
    DDV_MinMaxDouble(pDX, dl_shtock, dl_cylinder + 25 + 25, 1000);
    DDX_Text(pDX, IDC_EDIT2, dl_cylinder);
    DDV_MinMaxDouble(pDX, dl_cylinder, (h_nedokriska * 2) + 28, dl_shtock - 25);
    DDX_Text(pDX, IDC_EDIT3, h_ver_kriska);
    DDV_MinMaxDouble(pDX, h_ver_kriska, 15, dl_cylinder / 2);
    DDX_Text(pDX, IDC_EDIT4, h_nedokriska);
    DDV_MinMaxDouble(pDX, h_nedokriska, 25, dl_cylinder + 15);
}

BEGIN_MESSAGE_MAP(CSbDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CSbDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// Обработчики сообщений CSbDlg

BOOL CSbDlg::OnInitDialog()
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

void CSbDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CSbDlg::OnPaint()
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
HCURSOR CSbDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}
bool CSbDlg::CheckData()
{
	if (!UpdateData())
		return false;

	return true;
}



void CSbDlg::OnBnClickedButton1()
{
	// TODO: добавьте свой код обработчика уведомлений
	BeginWaitCursor();

	if (!CheckData())
		return;

	CComPtr<IUnknown> pKompasAppUnk = nullptr;
	if (!kompass)
	{
		// Получаем CLSID для Компас
		CLSID InvAppClsid;
		HRESULT hRes = CLSIDFromProgID(L"Kompas.Application.5", &InvAppClsid);
		if (FAILED(hRes)) {
			kompass = nullptr;
			return;
		}

		// Проверяем есть ли запущенный экземпляр Компас
		//если есть получаем IUnknown
		hRes = ::GetActiveObject(InvAppClsid, NULL, &pKompasAppUnk);
		if (FAILED(hRes)) {
			// Приходится запускать Компас самим так как работающего нет
			// Также получаем IUnknown для только что запущенного приложения Компас
			TRACE(L"Could not get hold of an active Inventor, will start a new session\n");
			hRes = CoCreateInstance(InvAppClsid, NULL, CLSCTX_LOCAL_SERVER, __uuidof(IUnknown), (void**)&pKompasAppUnk);
			if (FAILED(hRes)) {
				kompass = nullptr;
				return;
			}
		}

		// Получаем интерфейс приложения Компас
		hRes = pKompasAppUnk->QueryInterface(__uuidof(KompasObject), (void**)&kompass);
		if (FAILED(hRes)) {
			return;
		}
	}
	kompass->Visible = true;
	iDoc3D = kompass->Document3D();
	iDoc3D->Create(false, true);
	Part = iDoc3D->GetPart(pTop_Part);

    //цилиндр

    ksEntityPtr ISketchEntity = Part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef = ISketchEntity->GetDefinition();
    sketchDef->SetPlane(Part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity->Create();
    ksDocument2DPtr Doc2D = sketchDef->BeginEdit();

    Doc2D->ksCircle(0, 0, 45, 1);

    sketchDef->EndEdit();

    ksEntityPtr BossExtr = Part->NewEntity(o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef = BossExtr->GetDefinition();
    BossExtrDef->SetSideParam(TRUE, etBlind, dl_cylinder, 0, FALSE);
    BossExtrDef->directionType* (dtNormal);
    BossExtrDef->SetSketch(ISketchEntity);
    BossExtr->Create();


    ksEntityPtr ISketchEntity1 = Part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
    sketchDef1->SetPlane(Part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity1->Create();
    ksDocument2DPtr Doc2D1 = sketchDef1->BeginEdit();

    Doc2D1->ksCircle(0, 0, 40, 1);

    sketchDef1->EndEdit();

    ksEntityPtr CutExtr = Part->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef = CutExtr->GetDefinition();
    CutExtrDef->SetSideParam(FALSE, 1, dl_cylinder, 0, FALSE);
    CutExtrDef->SetSketch(ISketchEntity1);
    CutExtr->Create();

    ksEntityCollectionPtr flFaces = Part->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == CutExtr)
        {
            if (def->IsCylinder())
            {
                double h, r;
                def->GetCylinderParam(&h, &r);
                if (r == 40)
                {
                    face->Putname("circle_cylinder");
                    face->Update();
                }

            }
        }
    }
    flFaces = Part->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == BossExtr)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);

                    if (d->IsCircle())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;

                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 45 && z1 == dl_cylinder)
                        {
                            face->Putname("cylinder_plane1");
                            face->Update();
                            break;
                        }
                        if (x1 == 45 && z1 == 0)
                        {
                            face->Putname("cylinder_plane2");;
                            face->Update();
                            break;
                        }
                    }
                }
            }
        }
    }

    iDoc3D->SaveAs(L"C:\\с\\Part1.m3d");

    //шток

    ksDocument3DPtr iDoc3D1 = kompass->Document3D();
    iDoc3D1->Create(false, true);
    iDoc3D1 = kompass->ActiveDocument3D();

    ksPartPtr Part1;
    Part1 = iDoc3D1->GetPart(pTop_Part);

    ksEntityPtr ISketchEntity2 = Part1->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef2 = ISketchEntity2->GetDefinition();
    sketchDef2->SetPlane(Part1->GetDefaultEntity(o3d_planeXOZ));
    ISketchEntity2->Create();
    ksDocument2DPtr Doc2D2 = sketchDef2->BeginEdit();

    Doc2D2->ksLineSeg(-20, 0, 300, 0, 3);
    Doc2D2->ksLineSeg(0, 0, 0, 10.5, 1);
    Doc2D2->ksLineSeg(0, 10.5, 1.5, 12, 1);
    Doc2D2->ksLineSeg(1.5, 12, 20, 12, 1);
    Doc2D2->ksLineSeg(20, 12, 20.954, 11.0465, 1);
    Doc2D2->ksArcByAngle(21.307, 11.4, 0.5, 225, 270, 0, 1);
    Doc2D2->ksLineSeg(21.307, 10.9, 24.1, 10.9, 1);
    Doc2D2->ksArcByAngle(24.1, 11.9, 1, 270, 360, 0, 1);
    Doc2D2->ksLineSeg(25.1, 11.9, 25.1, 15.2, 1);
    Doc2D2->ksLineSeg(25.1, 15.2, 27.1, 16, 1);
    Doc2D2->ksLineSeg(27.1, 16, 27 + dl_shtock, 16, 1);
    Doc2D2->ksLineSeg(27 + dl_shtock, 16, 27 + dl_shtock, 15.5, 1);
    Doc2D2->ksLineSeg(27 + dl_shtock, 15.5, 35 + dl_shtock, 15.5, 1);
    Doc2D2->ksLineSeg(35 + dl_shtock, 15.5, 35 + dl_shtock, 9, 1);
    Doc2D2->ksLineSeg(35 + dl_shtock, 9, 42 + dl_shtock, 9, 1);
    Doc2D2->ksLineSeg(42 + dl_shtock, 9, 42 + dl_shtock, 16, 1);
    Doc2D2->ksLineSeg(42 + dl_shtock, 16, 44.5 + dl_shtock, 16, 1);
    Doc2D2->ksLineSeg(44.5 + dl_shtock, 16, 50 + dl_shtock, 12, 1);
    Doc2D2->ksLineSeg(50 + dl_shtock, 12, 67 + dl_shtock, 12, 1);
    Doc2D2->ksLineSeg(67 + dl_shtock, 12, 70 + dl_shtock, 9, 1);
    Doc2D2->ksLineSeg(70 + dl_shtock, 9, 70 + dl_shtock, 0, 1);

    sketchDef2->EndEdit();

    ksEntityPtr BossRot = Part1->NewEntity(o3d_bossRotated);
    ksBossRotatedDefinitionPtr BossRotDef = BossRot->GetDefinition();
    BossRotDef->SetSideParam(TRUE, 360);
    BossRotDef->directionType* (dtNormal);
    BossRotDef->SetSketch(ISketchEntity2);
    BossRot->Create();

    ksEntityPtr Offset = Part1->NewEntity(o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr OffsetDef = Offset->GetDefinition();
    OffsetDef->direction = true;
    OffsetDef->offset = -(70 + dl_shtock);
    OffsetDef->SetPlane(Part1->GetDefaultEntity(o3d_planeYOZ));
    Offset->Create();

    ksEntityPtr ISketchEntity3 = Part1->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef3 = ISketchEntity3->GetDefinition();
    sketchDef3->SetPlane(Offset);
    ISketchEntity3->Create();
    ksDocument2DPtr Doc2D3 = sketchDef3->BeginEdit();

    Doc2D3->ksLineSeg(-2.5, 9, 2.5, 9, 1);
    Doc2D3->ksLineSeg(2.5, 9, 2.5, 17.5, 1);
    Doc2D3->ksLineSeg(2.5, 17.5, -2.5, 17.5, 1);
    Doc2D3->ksLineSeg(-2.5, 17.5, -2.5, 9, 1);

    sketchDef3->EndEdit();

    ksEntityPtr CutExtr1 = Part1->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef1 = CutExtr1->GetDefinition();
    CutExtrDef1->SetSideParam(FALSE, 0, 25, 0, FALSE);
    CutExtrDef1->SetSketch(ISketchEntity3);
    CutExtr1->Create();

    flFaces = Part1->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == BossRot)
        {
            if (def->IsCylinder())
            {
                double h, r;
                def->GetCylinderParam(&h, &r);
                if (r == 16)
                {
                    face->Putname("circle_shtok");
                    face->Update();
                }
            }
        }
    }

    flFaces = Part1->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face1 = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face1->GetDefinition();
        if (def->GetOwnerEntity() == BossRot)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);
                    if (d->IsCircle())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;
                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 27+ dl_shtock && z1 == -16)
                        {
                            face1->Putname("shtok_plane1");
                            face1->Update();
                            break;
                        }
                    }
                }
            }
        }
    }
    flFaces = Part1->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face1 = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face1->GetDefinition();
        if (def->GetOwnerEntity() == BossRot)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);
                    if (d->IsCircle())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;
                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 70+ dl_shtock && y1 == -9)
                        {
                            face1->Putname("shtok_plane");
                            face1->Update();
                            break;
                        }
                    }
                }
            }
        }
    }

    iDoc3D1->SaveAs(L"C:\\с\\Part2.m3d");

    //крышка нижняя

    ksDocument3DPtr idoc3D = kompass->Document3D();
    idoc3D->Create(false, true);
    idoc3D = kompass->ActiveDocument3D();

    ksPartPtr part;
    part = idoc3D->GetPart(pTop_Part);

    ksEntityPtr ISketchEntitya = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDefa = ISketchEntitya->GetDefinition();
    sketchDefa->SetPlane(part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntitya->Create();
    ksDocument2DPtr doc2d = sketchDefa->BeginEdit();
    doc2d->ksLineSeg(-52.5, -40, -52.5, 40, 1);//левая прямая
    doc2d->ksLineSeg(-40, 52.5, 40, 52.5, 1);//верхняя прямая
    doc2d->ksLineSeg(52.5, 40, 52.5, -40, 1);//правая прямая
    doc2d->ksLineSeg(40, -52.5, -40, -52.5, 1);//нижняя прямая
    doc2d->ksArcByAngle(52.5 - 12.5, 52.5 - 12.5, 12.5, 0, 90, 1, 1);
    doc2d->ksArcByAngle(-(52.5 - 12.5), 52.5 - 12.5, 12.5, 90, 180, 1, 1);
    doc2d->ksArcByAngle(52.5 - 12.5, -(52.5 - 12.5), 12.5, 270, 0, 1, 1);
    doc2d->ksArcByAngle(-(52.5 - 12.5), -(52.5 - 12.5), 12.5, 180, 270, 1, 1);
    sketchDefa->EndEdit();

    ksEntityPtr BossExtra = part->NewEntity(o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDefa = BossExtra->GetDefinition();
    BossExtrDefa->SetSideParam(TRUE, etBlind, 25, 0, FALSE);
    BossExtrDefa->directionType* (dtNormal);
    BossExtrDefa->SetSketch(ISketchEntitya);
    BossExtra->Create();

    ksEntityPtr ISketchEntity1a = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef1a = ISketchEntity1a->GetDefinition();
    sketchDef1a->SetPlane(part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity1a->Create();
    ksDocument2DPtr doc2d1 = sketchDef1a->BeginEdit();

    doc2d1->ksCircle(0, 0, 35 / 2, 1);

    sketchDef1a->EndEdit();

    ksEntityPtr CutExtr1a = part->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef1a = CutExtr1a->GetDefinition();
    CutExtrDef1a->SetSideParam(FALSE, etThroughAll, 25, 0, FALSE);
    CutExtrDef1a->directionType* (dtNormal);
    CutExtrDef1a->SetSketch(ISketchEntity1a);
    CutExtr1a->Create();

    ksEntityPtr ISketchEntity2a = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef2a = ISketchEntity2a->GetDefinition();
    sketchDef2a->SetPlane(part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity2a->Create();
    ksDocument2DPtr doc2d2 = sketchDef2a->BeginEdit();

    doc2d2->ksCircle(0, 0, 49 / 2, 1);
    sketchDef2a->EndEdit();

    ksEntityPtr CutExtr2 = part->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef2 = CutExtr2->GetDefinition();
    CutExtrDef2->SetSideParam(TRUE, etBlind, 7, 0, FALSE);
    CutExtrDef2->directionType* (dtNormal);
    CutExtrDef2->SetSketch(ISketchEntity2a);
    CutExtr2->Create();

    ksEntityPtr ISketchEntity3a = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef3a = ISketchEntity3a->GetDefinition();
    sketchDef3a->SetPlane(part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity3a->Create();
    ksDocument2DPtr doc2d3 = sketchDef3a->BeginEdit();

    doc2d3->ksCircle(0, 0, 60 / 2, 1);
    doc2d3->ksCircle(0, 0, 35 / 2, 1);
    sketchDef3a->EndEdit();

    ksEntityPtr BossExtr2 = part->NewEntity(o3d_baseExtrusion);
    ksBaseExtrusionDefinitionPtr BossExtrDef2 = BossExtr2->GetDefinition();
    BossExtrDef2->SetSideParam(FALSE, etBlind, 5, 0, FALSE);
    BossExtrDef2->directionType = dtReverse;
    BossExtrDef2->SetSketch(ISketchEntity3a);
    BossExtr2->Create();

    ksEntityPtr ISketchEntity0 = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef0 = ISketchEntity0->GetDefinition();
    sketchDef0->SetPlane(part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity0->Create();

    ksEntityPtr Offset1 = part->NewEntity(o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr OffsetDef1 = Offset1->GetDefinition();
    OffsetDef1->direction = true;
    OffsetDef1->offset = 25;
    OffsetDef1->SetPlane(ISketchEntity0);
    Offset1->Create();

    ksEntityPtr ISketchEntity01 = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef01 = ISketchEntity01->GetDefinition();
    sketchDef01->SetPlane(Offset1);
    ISketchEntity01->Create();
    ksDocument2DPtr doc2d01 = sketchDef01->BeginEdit();

    doc2d01->ksCircle(0, 0, 60 / 2, 1);
    doc2d01->ksCircle(0, 0, 80 / 2, 1);
    sketchDef01->EndEdit();

    ksEntityPtr BossExtr3 = part->NewEntity(o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef3 = BossExtr3->GetDefinition();
    BossExtrDef3->SetSideParam(TRUE, etBlind, 15, 0, FALSE);
    BossExtrDef3->directionType* (dtNormal);
    BossExtrDef3->SetSketch(ISketchEntity01);
    BossExtr3->Create();

    ksEntityPtr ISketchEntity4 = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef4 = ISketchEntity4->GetDefinition();
    sketchDef4->SetPlane(part->GetDefaultEntity(o3d_planeXOZ));
    ISketchEntity4->Create();
    ksDocument2DPtr doc2d4 = sketchDef4->BeginEdit();

    doc2d4->ksLineSeg(-52.5, -12.5, -20.5, -12.5, 3);
    doc2d4->ksLineSeg(-52.5, -12.5, -52.5, -21.028, 1);
    doc2d4->ksLineSeg(-52.5, -21.028, -36.5, -19.899, 1);
    doc2d4->ksLineSeg(-36.5, -19.899, -36.5, -16.5, 1);
    doc2d4->ksLineSeg(-36.5, -16.5, -22.809, -16.5, 1);
    doc2d4->ksLineSeg(-22.809, -16.5, -20.5, -12.5, 1);
    sketchDef4->EndEdit();

    ksEntityPtr CutRot = part->NewEntity(o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef = CutRot->GetDefinition();
    CutRotDef->cut = TRUE;
    CutRotDef->SetSideParam(TRUE, 360);
    CutRotDef->directionType* (dtNormal);
    CutRotDef->SetSketch(ISketchEntity4);
    CutRot->Create();

    ksEntityPtr ISketchEntity5 = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef5 = ISketchEntity5->GetDefinition();
    sketchDef5->SetPlane(part->GetDefaultEntity(o3d_planeXOZ));
    ISketchEntity5->Create();
    ksDocument2DPtr doc2d5 = sketchDef5->BeginEdit();

    doc2d5->ksLineSeg(-26, 0 - 25, -26, 12.5 - 25, 3);
    doc2d5->ksLineSeg(-26, 12.5 - 25, -30, 8.5 - 25, 1);
    doc2d5->ksLineSeg(-30, 8.5 - 25, -30, 0 - 25, 1);
    doc2d5->ksLineSeg(-30, 0 - 25, -26, 0 - 25, 1);

    sketchDef5->EndEdit();

    ksEntityPtr CutRot1 = part->NewEntity(o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef1 = CutRot1->GetDefinition();
    CutRotDef1->cut = TRUE;
    CutRotDef1->SetSideParam(TRUE, 360);
    CutRotDef1->directionType* (dtNormal);
    CutRotDef1->SetSketch(ISketchEntity5);
    CutRot1->Create();

    ksEntityPtr ISketchEntity6 = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef6 = ISketchEntity6->GetDefinition();
    sketchDef6->SetPlane(part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity6->Create();
    ksDocument2DPtr doc2d6 = sketchDef6->BeginEdit();

    doc2d6->ksCircle(40, 40, 10, 1);

    sketchDef6->EndEdit();

    ksEntityPtr CutExtr3 = part->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef3 = CutExtr3->GetDefinition();
    CutExtrDef3->SetSideParam(TRUE, etBlind, 16, 0, FALSE);
    CutExtrDef3->directionType* (dtNormal);
    CutExtrDef3->SetSketch(ISketchEntity6);
    CutExtr3->Create();

    ksEntityPtr ISketchEntity7 = part->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef7 = ISketchEntity7->GetDefinition();
    sketchDef7->SetPlane(part->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity7->Create();
    ksDocument2DPtr doc2d7 = sketchDef7->BeginEdit();

    doc2d7->ksCircle(40, 40, 13 / 2, 1);

    sketchDef7->EndEdit();

    ksEntityPtr CutExtr4 = part->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef4 = CutExtr4->GetDefinition();
    CutExtrDef4->SetSideParam(false, etThroughAll, 25, 0, FALSE);
    CutExtrDef4->directionType* (dtNormal);
    CutExtrDef4->SetSketch(ISketchEntity7);
    CutExtr4->Create();

    ksEntityPtr MeshCopy = part->NewEntity(o3d_meshCopy);
    ksMeshCopyDefinitionPtr MeshCopyDef = MeshCopy->GetDefinition();
    MeshCopyDef->SetCopyParamAlongAxis(TRUE, -90, 2, 80, FALSE);
    MeshCopyDef->SetCopyParamAlongAxis(FALSE, -90, 2, 80, FALSE);
    ksEntityCollectionPtr EnColPartDef = MeshCopyDef->OperationArray();
    EnColPartDef->Add(CutExtr4);
    EnColPartDef->Add(CutExtr3);

    MeshCopy->Create();

    ksEntityCollectionPtr flFaces707 = part->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces707->GetCount(); i++)
    {
        ksEntityPtr face = flFaces707->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();

        //if (def->GetOwnerEntity() == BossExtr3)
       // {
        if (def->IsCylinder())
        {
            double h, r;
            def->GetCylinderParam(&h, &r);

            if (r == 40)
            {
                face->Putname("Cylinder_Nis_Kriska");
                face->Update();
            }
        }
        //}
    }

    ksEntityCollectionPtr flFaces20 = part->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces20->GetCount(); i++)
    {
        ksEntityPtr face = flFaces20->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == BossExtra)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);

                    if (d->IsArc())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;

                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 40 && z1 == 25)
                        {
                            face->Putname("Face_Niz_Kriska");
                            face->Update();
                            break;
                        }
                    }
                }
            }
        }
    }

    flFaces20 = part->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces20->GetCount(); i++)
    {
        ksEntityPtr face = flFaces20->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == BossExtr3)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);

                    if (d->IsCircle())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;

                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 40 && z1 == 40)
                        {
                            face->Putname("Face_for_Shtock");
                            face->Update();
                            break;
                        }
                    }
                }
            }
        }
    }
    idoc3D->SaveAs(L"C:\\с\\Part3.m3d");

    //недокрышка

    ksDocument3DPtr iDoc3D2 = kompass->Document3D();
    iDoc3D2->Create(false, true);
    iDoc3D2 = kompass->ActiveDocument3D();

    ksPartPtr Part2;
    Part2 = iDoc3D2->GetPart(pTop_Part);

    ksEntityPtr ISketchEnti = Part2->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchD = ISketchEnti->GetDefinition();
    sketchD->SetPlane(Part2->GetDefaultEntity(o3d_planeXOY));
    ISketchEnti->Create();
    ksDocument2DPtr Doc2D00 = sketchD->BeginEdit();

    Kompas6API5::ksMathematic2DPtr Mathematic2D;
    Mathematic2D = *new Kompas6API5::ksMathematic2DPtr(kompass->GetMathematic2D());

    Doc2D00->ksLineSeg(-(52.5 - 12.5), 52.5, 52.5 - 12.5, 52.5, 1);
    Doc2D00->ksLineSeg(52.5, 52.5 - 12.5, 52.5, -(52.5 - 12.5), 1);
    Doc2D00->ksLineSeg(52.5 - 12.5, -(52.5), -(52.5 - 12.5), -(52.5), 1);
    Doc2D00->ksLineSeg(-(52.5), -(52.5 - 12.5), -(52.5), 52.5 - 12.5, 1);

    Doc2D00->ksArcByAngle(52.5 - 12.5, 52.5 - 12.5, 12.5, 0, 90, 1, 1);
    Doc2D00->ksArcByAngle(-(52.5 - 12.5), 52.5 - 12.5, 12.5, 90, 180, 1, 1);
    Doc2D00->ksArcByAngle(52.5 - 12.5, -(52.5 - 12.5), 12.5, 270, 0, 1, 1);
    Doc2D00->ksArcByAngle(-(52.5 - 12.5), -(52.5 - 12.5), 12.5, 180, 270, 1, 1);
    sketchD->EndEdit();

    ksEntityPtr BossEx = Part2->NewEntity(o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrD = BossEx->GetDefinition();
    BossExtrD->SetSideParam(TRUE, etBlind, h_nedokriska, 0, FALSE);
    BossExtrD->directionType* (dtNormal);
    BossExtrD->SetSketch(ISketchEnti);
    BossEx->Create();

    ksEntityPtr ISketchEnti1 = Part2->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchD1 = ISketchEnti1->GetDefinition();
    sketchD1->SetPlane(Part2->GetDefaultEntity(o3d_planeXOY));
    ISketchEnti1->Create();
    ksDocument2DPtr Doc2D001 = sketchD1->BeginEdit();

    Doc2D001->ksCircle(0, 0, 45, 1);

    sketchD1->EndEdit();

    ksEntityPtr CutEx = Part2->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrD = CutEx->GetDefinition();
    CutExtrD->cut* (TRUE);
    CutExtrD->directionType* (dtNormal);
    CutExtrD->SetSideParam(FALSE, etThroughAll, h_nedokriska, 0, FALSE);
    CutExtrD->SetSketch(ISketchEnti1);
    CutEx->Create();

    ksEntityPtr ISketchEntit0 = Part2->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDe0 = ISketchEntit0->GetDefinition();
    sketchDe0->SetPlane(Part2->GetDefaultEntity(o3d_planeXOZ));
    ISketchEntit0->Create();

    ksEntityPtr Offset2 = Part2->NewEntity(o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr OffsetDef2 = Offset2->GetDefinition();
    OffsetDef2->direction = true;
    OffsetDef2->offset = 52.5 - 12.5;
    OffsetDef2->SetPlane(ISketchEntit0);
    Offset2->Create();

    ksEntityPtr ISketchEntit01 = Part2->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDe01 = ISketchEntit01->GetDefinition();
    sketchDe01->SetPlane(Offset2);
    ISketchEntit01->Create();
    ksDocument2DPtr Doc2D01 = sketchDe01->BeginEdit();

    Doc2D01->ksLineSeg(40, 0, 40, -h_nedokriska, 1);
    Doc2D01->ksLineSeg(40, -h_nedokriska, 40 + 6, -h_nedokriska, 1);
    Doc2D01->ksLineSeg(40 + 6, -h_nedokriska, 40 + 6 - 1.6, -(h_nedokriska - 1.6), 1);
    Doc2D01->ksLineSeg(40 + 6 - 1.6, -(h_nedokriska - 1.6), 40 + 6 - 1.6, 0, 1);
    Doc2D01->ksLineSeg(40 + 6 - 1.6, 0, 40, 0, 1);
    Doc2D01->ksLineSeg(40, 0, 40, -h_nedokriska, 3);
    sketchDe01->EndEdit();

    ksEntityPtr CutRExtr1 = Part2->NewEntity(o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRExtrDef1 = CutRExtr1->GetDefinition();
    CutRExtrDef1->cut* (TRUE);
    CutRExtrDef1->directionType* (dtNormal);
    CutRExtrDef1->SetSideParam(TRUE, 360);
    CutRExtrDef1->SetSketch(ISketchEntit01);
    CutRExtr1->Create();

    ksEntityPtr CircArr = Part2->NewEntity(o3d_circularCopy);
    ksCircularCopyDefinitionPtr CircDef = CircArr->GetDefinition();
    CircDef->SetAxis(Part2->GetDefaultEntity(o3d_axisOZ));
    CircDef->SetCopyParamAlongDir(4, 360, true, false);
    ksEntityCollectionPtr ColEn = CircDef->GetOperationArray();
    ColEn->Add(CutRExtr1);
    CircArr->Create();

    flFaces = Part2->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == CutEx)
        {
            if (def->IsCylinder())
            {
                double h, r;
                def->GetCylinderParam(&h, &r);

                if (r == 45)
                {
                    face->Putname("circle_inside");
                    face->Update();
                }
            }
        }
    }
    flFaces = Part2->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == BossEx)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);
                    if (d->IsArc())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;
                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 52.5 && z1 == h_nedokriska && y1 == 40)
                        {
                            face->Putname("plane1_nedokrishka");
                            face->Update();
                            break;
                        }
                        if (x1 == 52.5 && z1 == 0 && y1 == 40)
                        {
                            face->Putname("plane2_nedokrishka");
                            face->Update();
                            break;
                        }
                    }
                }
            }
        }
    }

    iDoc3D2->SaveAs(L"C:\\с\\Part4.m3d");

    //крышка верхняя

    ksDocument3DPtr iDoc3d = kompass->Document3D();
    iDoc3d->Create(false, true);
    iDoc3d = kompass->ActiveDocument3D();

    ksPartPtr Part3;
    Part3 = iDoc3d->GetPart(pTop_Part);

    ksEntityPtr ISketchEnti0 = Part3->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchD0 = ISketchEnti0->GetDefinition();
    sketchD0->SetPlane(Part3->GetDefaultEntity(o3d_planeXOY));
    ISketchEnti0->Create();
    ksDocument2DPtr doc2d001 = sketchD0->BeginEdit();

    doc2d001->ksArcByPoint(12.5, 12.5, 12.5, 0, 12.5, 12.5, 0, 2, 1);
    doc2d001->ksLineSeg(12.5, 0, 92.5, 0, 1);
    doc2d001->ksArcByPoint(92.5, 12.5, 12.5, 92.5, 0, 105, 12.5, 1, 1);
    doc2d001->ksLineSeg(105, 12.5, 105, 92.5, 1);
    doc2d001->ksArcByPoint(92.5, 92.5, 12.5, 105, 92.5, 92.5, 105, 1, 1);
    doc2d001->ksLineSeg(92.5, 105, 12.5, 105, 1);
    doc2d001->ksArcByPoint(12.5, 92.5, 12.5, 12.5, 105, 0, 92.5, 1, 1);
    doc2d001->ksLineSeg(0, 92.5, 0, 12.5, 1);

    sketchD0->EndEdit();

    ksEntityPtr BossExtr03 = Part3->NewEntity(o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef03 = BossExtr03->GetDefinition();
    BossExtrDef03->SetSideParam(TRUE, etBlind, 26, 0, FALSE);
    BossExtrDef03->directionType* (dtNormal);
    BossExtrDef03->SetSketch(ISketchEnti0);
    BossExtr03->Create();

    ksEntityPtr EnPlaneOffset = Part3->NewEntity(o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef = EnPlaneOffset->GetDefinition();
    PlaneOffsetDef->direction = FALSE;
    PlaneOffsetDef->offset = 12.5;
    PlaneOffsetDef->SetPlane(Part3->GetDefaultEntity(o3d_planeYOZ));
    EnPlaneOffset->Create();

    ksEntityPtr ISketchEntity001 = Part3->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef001 = ISketchEntity001->GetDefinition();
    sketchDef001->SetPlane(EnPlaneOffset);
    ISketchEntity001->Create();
    ksDocument2DPtr doc2d002 = sketchDef001->BeginEdit();

    doc2d002->ksLineSeg(-26, -2.5, -26, -12.5, 1);
    doc2d002->ksLineSeg(-26, -12.5, 0, -12.5, 3);
    doc2d002->ksLineSeg(0, -12.5, 0, -6, 1);
    doc2d002->ksLineSeg(0, -6, -10, -6, 1);
    doc2d002->ksLineSeg(-10, -6, -10, -2.5, 1);
    doc2d002->ksLineSeg(-10, -2.5, -26, -2.5, 1);

    sketchDef001->EndEdit();

    ksEntityPtr EnCutRot01 = Part3->NewEntity(o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef01 = EnCutRot01->GetDefinition();
    CutRotDef01->directionType* (dtNormal);
    CutRotDef01->cut = TRUE;
    CutRotDef01->SetSideParam(TRUE, 360);
    CutRotDef01->SetSketch(ISketchEntity001);
    EnCutRot01->Create();

    ksEntityPtr EnAxis = Part3->NewEntity(o3d_axisOX);
    EnAxis->Create();

    ksEntityPtr MeshCopy1 = Part3->NewEntity(o3d_meshCopy);
    ksMeshCopyDefinitionPtr MeshCopyDef1 = MeshCopy1->GetDefinition();
    MeshCopyDef1->insideFlag = FALSE;
    MeshCopyDef1->SetCopyParamAlongAxis(TRUE, 0, 2, 80, FALSE);
    MeshCopyDef1->SetAxis2(EnAxis);
    MeshCopyDef1->SetCopyParamAlongAxis(FALSE, 90, 2, 80, FALSE);
    ksEntityCollectionPtr EnColPartDef1 = MeshCopyDef1->OperationArray();
    EnColPartDef1->Clear();
    EnColPartDef1->Add(EnCutRot01);
    MeshCopy1->Create();

    ksEntityPtr ISketchEntity03 = Part3->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef03 = ISketchEntity03->GetDefinition();
    sketchDef03->SetPlane(Part3->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity03->Create();

    ksDocument2DPtr doc2d03 = sketchDef03->BeginEdit();
    doc2d03->ksCircle(52.5, 52.5, 30, 1);
    doc2d03->ksCircle(52.5, 52.5, 40, 1);
    sketchDef03->EndEdit();

    ksEntityPtr BossExtr4 = Part3->NewEntity(o3d_baseExtrusion);
    ksBaseExtrusionDefinitionPtr BossExtrDef4 = BossExtr4->GetDefinition();
    BossExtrDef4->SetSideParam(FALSE, etBlind, h_ver_kriska, 0, FALSE);
    BossExtrDef4->directionType = dtReverse;
    BossExtrDef4->SetSketch(ISketchEntity03);
    BossExtr4->Create();

    ksEntityPtr EnPlaneOffset1 = Part3->NewEntity(o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef1 = EnPlaneOffset1->GetDefinition();
    PlaneOffsetDef1->direction = FALSE;
    PlaneOffsetDef1->offset = 62.5;
    PlaneOffsetDef1->SetPlane(Part3->GetDefaultEntity(o3d_planeYOZ));
    EnPlaneOffset1->Create();

    ksEntityPtr ISketchEntity04 = Part3->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef04 = ISketchEntity04->GetDefinition();
    sketchDef04->SetPlane(EnPlaneOffset1);
    ISketchEntity04->Create();
    ksDocument2DPtr doc2d04 = sketchDef04->BeginEdit();

    doc2d04->ksLineSeg(-26, -18.5, -31, -18.5, 1);
    doc2d04->ksLineSeg(-31, -18.5, -68, -52.5, 1);
    doc2d04->ksLineSeg(-68, -52.5, -58, -52.5, 1);
    doc2d04->ksArcBy3Points(-58, -52.5, -50, -44.5, -42, -52.5, 1);
    doc2d04->ksLineSeg(-42, -52.5, -26, -52.5, 1);
    doc2d04->ksLineSeg(-26, -52.5, -26, -18.5, 1);

    sketchDef04->EndEdit();

    ksEntityPtr BossExtr04 = Part3->NewEntity(o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef04 = BossExtr04->GetDefinition();
    BossExtrDef04->SetSideParam(TRUE, etBlind, 20, 0, FALSE);
    BossExtrDef04->directionType* (dtNormal);
    BossExtrDef04->SetSketch(ISketchEntity04);
    BossExtr04->Create();

    ksEntityPtr EnPlaneOffset2 = Part3->NewEntity(o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef2 = EnPlaneOffset2->GetDefinition();
    PlaneOffsetDef2->direction = TRUE;
    PlaneOffsetDef2->offset = 52.5;
    PlaneOffsetDef2->SetPlane(Part3->GetDefaultEntity(o3d_planeXOZ));
    EnPlaneOffset2->Create();

    ksEntityPtr EnMirrorOp = Part3->NewEntity(o3d_mirrorOperation);
    ksMirrorCopyDefinitionPtr MirrorCopyDef = EnMirrorOp->GetDefinition();
    MirrorCopyDef->SetPlane(EnPlaneOffset2);
    ksEntityCollectionPtr EnColl = MirrorCopyDef->GetOperationArray();
    EnColl->Clear();
    EnColl->Add(BossExtr04);
    EnMirrorOp->Create();

    ksEntityPtr EnPlaneOffset5 = Part3->NewEntity(o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef5 = EnPlaneOffset5->GetDefinition();
    PlaneOffsetDef5->direction = FALSE;
    PlaneOffsetDef5->offset = 52.5;
    PlaneOffsetDef5->SetPlane(Part3->GetDefaultEntity(o3d_planeYOZ));
    EnPlaneOffset5->Create();

    ksEntityPtr ISketchEntity05 = Part3->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef05 = ISketchEntity05->GetDefinition();
    sketchDef05->SetPlane(EnPlaneOffset5);
    ISketchEntity05->Create();
    ksDocument2DPtr doc2d05 = sketchDef05->BeginEdit();

    doc2d05->ksLineSeg(-13, 0, -21.53, 0, 1);
    doc2d05->ksLineSeg(-21.53, 0, -20.04, -16, 1);
    doc2d05->ksLineSeg(-20.04, -16, -17, -16, 1);
    doc2d05->ksLineSeg(-17, -16, -17, -32, 1);
    doc2d05->ksLineSeg(-17, -32, -13, -34.31, 1);
    doc2d05->ksLineSeg(-13, -34.31, -13, 0, 3);
    sketchDef05->EndEdit();

    ksEntityPtr EnCutRot5 = Part3->NewEntity(o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef5 = EnCutRot5->GetDefinition();
    CutRotDef5->directionType* (dtNormal);
    CutRotDef5->cut = TRUE;
    CutRotDef5->SetSideParam(TRUE, 360);
    CutRotDef5->SetSketch(ISketchEntity05);
    EnCutRot5->Create();

    ksEntityPtr ISketchEntity06 = Part3->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef06 = ISketchEntity06->GetDefinition();
    sketchDef06->SetPlane(EnPlaneOffset5);
    ISketchEntity06->Create();
    ksDocument2DPtr doc2d06 = sketchDef06->BeginEdit();

    doc2d06->ksLineSeg(0, -22.5, -10, -22.5, 1);
    doc2d06->ksLineSeg(-10, -22.5, -10, -26.5, 1);
    doc2d06->ksLineSeg(-10, -26.5, 0, -26.5, 3);
    doc2d06->ksLineSeg(0, -26.5, 0, -22.5, 1);
    sketchDef06->EndEdit();

    ksEntityPtr EnCutRot6 = Part3->NewEntity(o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef6 = EnCutRot6->GetDefinition();
    CutRotDef6->directionType* (dtNormal);
    CutRotDef6->cut = TRUE;
    CutRotDef6->SetSideParam(TRUE, 360);
    CutRotDef6->SetSketch(ISketchEntity06);
    EnCutRot6->Create();

    ksEntityPtr Fillet = Part3->NewEntity(o3d_fillet);
    ksFilletDefinitionPtr FilletDef = Fillet->GetDefinition();
    FilletDef->radius = 3;
    FilletDef->tangent = FALSE;

    ksEntityCollectionPtr EnColPart = Part3->EntityCollection(o3d_edge);
    ksEntityCollectionPtr EnColFillet = FilletDef->array();
    EnColFillet->Clear();

    EnColPart->SelectByPoint(52.5, 82.5, 0);
    EnColFillet->Add(EnColPart->GetByIndex(0));
    Fillet->Create();

    ksEntityPtr Fillet2 = Part3->NewEntity(o3d_fillet);
    ksFilletDefinitionPtr FilletDef2 = Fillet2->GetDefinition();
    FilletDef2->radius = 8;
    FilletDef2->tangent = FALSE;

    ksEntityCollectionPtr EnColPart2 = Part3->EntityCollection(o3d_edge);
    ksEntityCollectionPtr EnColFillet2 = FilletDef2->array();
    EnColFillet2->Clear();

    EnColPart2->SelectByPoint(52.5, 52.5, 68);
    EnColFillet2->Add(EnColPart2->GetByIndex(0));
    Fillet2->Create();

    ksEntityCollectionPtr flFaces3 = Part3->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces3->GetCount(); i++)
    {
        ksEntityPtr face = flFaces3->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();

        //if (def->GetOwnerEntity() == BossExtr3)
        // {
        if (def->IsCylinder())
        {
            double h, r;
            def->GetCylinderParam(&h, &r);

            if (r == 40)
            {
                face->Putname("Cylinder_Ver_Kriska");
                face->Update();
            }
        }
        //}
    }
    ksEntityCollectionPtr flFaces31 = Part3->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces31->GetCount(); i++)
    {
        ksEntityPtr face = flFaces31->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        if (def->GetOwnerEntity() == BossExtr03)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);

                    if (d->IsArc())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;

                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 12.5 && z1 == 0)
                        {
                            face->Putname("Face_Ver_Kriska");
                            face->Update();
                            break;
                        }
                    }
                }
            }
        }
    }

    iDoc3d->SaveAs(L"C:\\с\\Part5.m3d");

    //поршень

    ksDocument3DPtr iDocument3D = kompass->Document3D();
    iDocument3D->Create(false, true);
    iDocument3D = kompass->ActiveDocument3D();

    ksPartPtr Part4;
    Part4 = iDocument3D->GetPart(pTop_Part);

    ksEntityPtr ISketchEntity02 = Part4->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef02 = ISketchEntity02->GetDefinition();
    sketchDef02->SetPlane(Part4->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity02->Create();
    ksDocument2DPtr dokument2D = sketchDef02->BeginEdit();

    dokument2D->ksLineSeg(0, 0, 0, 40, 3);
    dokument2D->ksLineSeg(16, 0, 36.535898, 0, 1);
    dokument2D->ksLineSeg(36.535898, 0, 40, 2, 1);
    dokument2D->ksLineSeg(40, 2, 40, 4, 1);
    dokument2D->ksLineSeg(40, 4, 32, 4, 1);
    dokument2D->ksLineSeg(32, 4, 32, 11.5, 1);
    dokument2D->ksLineSeg(32, 11.5, 40, 11.5, 1);
    dokument2D->ksLineSeg(40, 11.5, 40, 16.5, 1);
    dokument2D->ksLineSeg(40, 16.5, 32, 16.5, 1);
    dokument2D->ksLineSeg(32, 16.5, 32, 24, 1);
    dokument2D->ksLineSeg(32, 24, 40, 24, 1);
    dokument2D->ksLineSeg(40, 24, 40, 26, 1);
    dokument2D->ksLineSeg(40, 26, 36.535898, 28, 1);
    dokument2D->ksLineSeg(36.535898, 28, 16, 28, 1);
    dokument2D->ksLineSeg(16, 28, 16, 0, 1);
    sketchDef02->EndEdit();

    ksEntityPtr EntityRotated = Part4->NewEntity(o3d_baseRotated);
    ksBaseRotatedDefinitionPtr BaseRotatedDef = EntityRotated->GetDefinition();
    BaseRotatedDef->SetSideParam(TRUE, 360);
    BaseRotatedDef->directionType* (dtNormal);
    BaseRotatedDef->SetSketch(ISketchEntity02);
    EntityRotated->Create();

    flFaces = Part4->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();

        /* if (def->GetOwnerEntity() == EntityRotated)*/
        if (def->IsCylinder())
        {
            double h, r;
            def->GetCylinderParam(&h, &r);
            if (r == 16)
            {
                face->Putname("Cylinder_Porshen");
                face->Update();
            }
        }
    }
    flFaces = Part4->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces->GetCount(); i++)
    {
        ksEntityPtr face = flFaces->GetByIndex(i);
        ksFaceDefinitionPtr def = face->GetDefinition();
        //if (def->GetOwnerEntity() == EntityRotated)
        //{
            if (def->IsPlanar())
            {
                /*double h, r;
                def->GetCylinderParam(&h, &r);
                if (r == 16)
                {*/
                    ksEdgeCollectionPtr col = def->EdgeCollection();
                    for (int k = 0; k < col->GetCount(); k++)
                    {
                        ksEdgeDefinitionPtr d = col->GetByIndex(k);
                        if (d->IsCircle())
                        {
                            ksVertexDefinitionPtr p1 = d->GetVertex(true);
                            ksVertexDefinitionPtr p2 = d->GetVertex(false);
                            double x1, y1, z1;
                            p1->GetPoint(&x1, &y1, &z1);
                            if (x1 == 16 && y1 == 0)
                            {
                                face->Putname("Face_Porshen");
                                face->Update();
                                break;
                            }
                        }
                    }
                //}
            }
        //}
    }

    iDocument3D->SaveAs(L"C:\\с\\Part6.m3d");

    //прокладка

    ksDocument3DPtr iDocument3D1 = kompass->Document3D();
    iDocument3D1->Create(false, true);
    iDocument3D1 = kompass->ActiveDocument3D();

    ksPartPtr Part5;
    Part5 = iDocument3D1->GetPart(pTop_Part);

    ksEntityPtr ISketchEntity007 = Part5->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef007 = ISketchEntity007->GetDefinition();
    sketchDef007->SetPlane(Part5->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity007->Create();
    ksDocument2DPtr doc2d007 = sketchDef007->BeginEdit();

    doc2d007->ksArcByPoint(40, 40, 12.5, 40, 52.5, 52.5, 40, -1, 1);
    doc2d007->ksArcByPoint(-40, 40, 12.5, -52.5, 40, -40, 52.5, -1, 1);
    doc2d007->ksArcByPoint(-40, -40, 12.5, -52.5, -40, -40, -52.5, 1, 1);
    doc2d007->ksArcByPoint(40, -40, 12.5, 40, -52.5, 52.5, -40, 1, 1);

    doc2d007->ksLineSeg(-40, 52.5, 40, 52.5, 1);
    doc2d007->ksLineSeg(-52.5, 40, -52.5, -40, 1);
    doc2d007->ksLineSeg(-40, -52.5, 40, -52.5, 1);
    doc2d007->ksLineSeg(52.5, -40, 52.5, 40, 1);

    sketchDef007->EndEdit();

    ksEntityPtr BossExtr5 = Part5->NewEntity(o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef5 = BossExtr5->GetDefinition();
    BossExtrDef5->SetSideParam(TRUE, etBlind, 3, 0, FALSE);
    BossExtrDef5->directionType* (dtNormal);
    BossExtrDef5->SetSketch(ISketchEntity007);
    BossExtr5->Create();

    ksEntityPtr ISketchEntity8 = Part5->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef8 = ISketchEntity8->GetDefinition();
    sketchDef8->SetPlane(Part5->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity8->Create();
    ksDocument2DPtr doc2d8 = sketchDef8->BeginEdit();

    doc2d8->ksCircle(0, 0, 40, 1);

    sketchDef8->EndEdit();

    ksEntityPtr CutExtr5 = Part5->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef5 = CutExtr5->GetDefinition();
    CutExtrDef5->SetSideParam(TRUE, etThroughAll, 3, 0, FALSE);
    CutExtrDef5->directionType* (dtNormal);
    CutExtrDef5->SetSketch(ISketchEntity8);
    CutExtr5->Create();

    ksEntityPtr ISketchEntity9 = Part5->NewEntity(o3d_sketch);
    ksSketchDefinitionPtr sketchDef9 = ISketchEntity9->GetDefinition();
    sketchDef9->SetPlane(Part5->GetDefaultEntity(o3d_planeXOY));
    ISketchEntity9->Create();
    ksDocument2DPtr doc2d9 = sketchDef9->BeginEdit();

    doc2d9->ksCircle(40, 40, 5, 1);
    doc2d9->ksCircle(-40, 40, 5, 1);
    doc2d9->ksCircle(-40, -40, 5, 1);
    doc2d9->ksCircle(40, -40, 5, 1);

    sketchDef9->EndEdit();

    ksEntityPtr CutExtr9 = Part5->NewEntity(o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef9 = CutExtr9->GetDefinition();
    CutExtrDef9->SetSideParam(TRUE, etThroughAll, 3, 0, FALSE);
    CutExtrDef9->directionType* (dtNormal);
    CutExtrDef9->SetSketch(ISketchEntity9);
    CutExtr9->Create();

    ksEntityCollectionPtr flFaces5 = Part5->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces5->GetCount(); i++)
    {
        ksEntityPtr face5 = flFaces5->GetByIndex(i);
        ksFaceDefinitionPtr def5 = face5->GetDefinition();
        if (def5->IsCylinder())
        {
            double h, r;
            def5->GetCylinderParam(&h, &r);
            if (r == 40)
            {
                face5->Putname("Cylinder_Prokladka");
                face5->Update();
            }
        }
    }
    ksEntityCollectionPtr flFaces6 = Part5->EntityCollection(o3d_face);
    flFaces6 = Part5->EntityCollection(o3d_face);
    for (int i = 0; i < flFaces6->GetCount(); i++)
    {
        ksEntityPtr face6 = flFaces6->GetByIndex(i);
        ksFaceDefinitionPtr def = face6->GetDefinition();
        if (def->GetOwnerEntity() == BossExtr5)
        {
            if (def->IsPlanar())
            {
                ksEdgeCollectionPtr col = def->EdgeCollection();
                for (int k = 0; k < col->GetCount(); k++)
                {
                    ksEdgeDefinitionPtr d = col->GetByIndex(k);
                    if (d->IsArc())
                    {
                        ksVertexDefinitionPtr p1 = d->GetVertex(true);
                        double x1, y1, z1;
                        p1->GetPoint(&x1, &y1, &z1);
                        if (x1 == 40 && z1 == 0)
                        {
                            face6->Putname("Plane1_prokladka");
                            face6->Update();
                            break;
                        }
                        if (x1 == 40 && z1 == 3)
                        {
                            face6->Putname("Plane2_prokladka");
                            face6->Update();
                            break;
                        }
                    }
                }
            }
        }
    }

    iDocument3D1->SaveAs(L"C:\\с\\Part7.m3d");


    ////добавление деталей

    iDoc3Dsb = kompass->Document3D();
    iDoc3Dsb->Create(false, false);
    iDoc3Dsb = kompass->ActiveDocument3D();
    Partsb = iDoc3Dsb->GetPart(pTop_Part);

    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part2.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part1.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part3.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part4.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part4.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part5.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part6.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part7.m3d", Partsb, VARIANT_FALSE);
    iDoc3Dsb->SetPartFromFile(L"C:\\с\\Part7.m3d", Partsb, VARIANT_FALSE);

    ksPartCollectionPtr partCollect = iDoc3Dsb->PartCollection(VARIANT_TRUE);
    int x;
    x = partCollect->GetCount();
    //шток
    Partsb = partCollect->GetByIndex(0);
    ksEntityCollectionPtr PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr circle_shtok = PartCol->GetByName("circle_shtok", true, true);
    ksEntityPtr shtok_plane1 = PartCol->GetByName("shtok_plane1", true, true);
    ksEntityPtr shtok_plane = PartCol->GetByName("shtok_plane", true, true);

    //цилиндр
    Partsb = partCollect->GetByIndex(1);
    PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr circle_cylinder = PartCol->GetByName("circle_cylinder", true, true);
    ksEntityPtr cylinder_plane1 = PartCol->GetByName("cylinder_plane1", true, true);
    ksEntityPtr cylinder_plane2 = PartCol->GetByName("cylinder_plane2", true, true);

    //поршень
    Partsb = partCollect->GetByIndex(6);
    PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr circle_porshen = PartCol->GetByName("Cylinder_Porshen", true, true);
    ksEntityPtr plane1_porshen = PartCol->GetByName("Face_Porshen", true, true);

    //крышка верхняя
    Partsb = partCollect->GetByIndex(5);
    PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr Cylinder_Ver_Kriska = PartCol->GetByName("Cylinder_Ver_Kriska", true, true);
    ksEntityPtr Face_Ver_Kriska = PartCol->GetByName("Face_Ver_Kriska", true, true);

    //Прокладка1
    Partsb = partCollect->GetByIndex(7);
    PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr Cylinder_Prokladka = PartCol->GetByName("Cylinder_Prokladka", true, true);
    ksEntityPtr Plane1_prokladka = PartCol->GetByName("Plane1_prokladka", true, true);
    ksEntityPtr Plane2_prokladka = PartCol->GetByName("Plane2_prokladka", true, true);

    //Прокладка2
    Partsb = partCollect->GetByIndex(8);
    PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr Cylinder_Prokladka1 = PartCol->GetByName("Cylinder_Prokladka", true, true);
    ksEntityPtr Plane1_prokladka1 = PartCol->GetByName("Plane1_prokladka", true, true);
    ksEntityPtr Plane2_prokladka1 = PartCol->GetByName("Plane2_prokladka", true, true);

    //недокрышка1
    Partsb = partCollect->GetByIndex(3);
    PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr circle_inside = PartCol->GetByName("circle_inside", true, true);
    ksEntityPtr plane_krishka_2piece1 = PartCol->GetByName("plane1_nedokrishka", true, true);


    //недокрышка2
    Partsb = partCollect->GetByIndex(4);
    PartCol = Partsb->EntityCollection(o3d_face);
    ksEntityPtr circle_inside1 = PartCol->GetByName("circle_inside", true, true);
    ksEntityPtr plane_krishka_2piece2 = PartCol->GetByName("plane2_nedokrishka", true, true);

    //крышка нижняя
    Partsb = partCollect->GetByIndex(2);
    PartCol = Partsb->EntityCollection(o3d_face);

    ksEntityPtr Cylinder_Nis_Kriska = PartCol->GetByName("Cylinder_Nis_Kriska", true, true);
    ksEntityPtr krishka_plane1 = PartCol->GetByName("Face_Niz_Kriska", true, true);
    ksEntityPtr krishka_niz = PartCol->GetByName("Face_for_Shtock", true, true);


    //цилинлр шток
    iDoc3Dsb->AddMateConstraint(mc_Concentric, circle_shtok, circle_cylinder, 0, 0, 0);

    //шток и поршень
    iDoc3Dsb->AddMateConstraint(mc_Concentric, circle_shtok, circle_porshen, 0, 1, 0);
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, shtok_plane1, plane1_porshen, 0, 1, 0);

    //шток и крышка верхняя
    iDoc3Dsb->AddMateConstraint(mc_Concentric, circle_shtok, Cylinder_Ver_Kriska, 0, 1, 0);

    //крышка верхняя прокладка1
    iDoc3Dsb->AddMateConstraint(mc_Concentric, Cylinder_Ver_Kriska, Cylinder_Prokladka, 0, 1, 0);
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, Face_Ver_Kriska, Plane1_prokladka, 0, 1, 0);

    //цининдр прокладка1
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, Plane2_prokladka, cylinder_plane1, 0, 1, 0);

    //шток недокрышка1
    iDoc3Dsb->AddMateConstraint(mc_Concentric, circle_inside, circle_shtok, 0, 1, 0);
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, plane_krishka_2piece1, Plane2_prokladka, 0, 1, 0);

    //шток недокрышка2
    iDoc3Dsb->AddMateConstraint(mc_Concentric, circle_shtok, circle_inside1, 0, 1, 0);

    //прокладка2 цилиндр
    iDoc3Dsb->AddMateConstraint(mc_Concentric, Cylinder_Prokladka1, circle_shtok, 0, 1, 0);
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, Plane1_prokladka1, cylinder_plane2, 0, 1, 0);

    //недокрышка2 прокладка2
    iDoc3Dsb->AddMateConstraint(mc_Concentric, Cylinder_Prokladka1, circle_inside1, -1, 1, 0);
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, plane_krishka_2piece2, Plane1_prokladka1, -1, 2, 0);

    //крышка нижняя шток
    iDoc3Dsb->AddMateConstraint(mc_Concentric, circle_shtok, Cylinder_Nis_Kriska, 0, 1, 0);
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, Plane2_prokladka1, krishka_plane1, 0, 1, 0);
    iDoc3Dsb->AddMateConstraint(mc_Coincidence, Face_Ver_Kriska, shtok_plane, 0, 1, 0);


    kompass.Detach();
}
