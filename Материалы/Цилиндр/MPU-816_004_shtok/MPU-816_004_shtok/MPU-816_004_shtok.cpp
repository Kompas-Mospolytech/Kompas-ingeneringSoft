//*********************************************************************************
// MPU-816_004_shtok
// Автор: Конинина А. Д.
// Дата создания: 10.11.2022
// Дата последнего изменения: 10.11.2022
//*********************************************************************************


#include <iostream>
#import "D:\\Bin\\kAPI5.tlb"
#import "D:\\Bin\\ksConstants3D.tlb"

using namespace std;
using namespace Kompas6API5;


int main()
{
    CoInitialize(NULL);
    KompasObjectPtr kompass;
    HRESULT hRes = kompass.GetActiveObject(L"Kompas.Application.5");
    if (FAILED(hRes))
        kompass.CreateInstance(L"Kompas.Application.5");
    kompass->Visible = true;

    ksDocument3DPtr iDoc3D = kompass->Document3D();
    iDoc3D->Create(false, true);
    iDoc3D = kompass->ActiveDocument3D();

    ksPartPtr Part;
    Part = iDoc3D->GetPart(Kompas6Constants3D::pTop_Part);

    ksEntityPtr ISketchEntity = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef = ISketchEntity->GetDefinition();
    sketchDef->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOZ));
    ISketchEntity->Create();
    ksDocument2DPtr Doc2D = sketchDef->BeginEdit();

    Doc2D->ksLineSeg(-20, 0, 300, 0, 3); 

    Doc2D->ksLineSeg(0, 0, 0, 10.5, 1);
    Doc2D->ksLineSeg(0, 10.5, 1.5, 12, 1);
    Doc2D->ksLineSeg(1.5, 12, 20, 12, 1);
    Doc2D->ksLineSeg(20, 12, 20.954, 11.0465, 1);
    Doc2D->ksArcByAngle(21.307, 11.4, 0.5, 225, 270, 0, 1);
    Doc2D->ksLineSeg(21.307, 10.9, 24.1, 10.9, 1);
    Doc2D->ksArcByAngle(24.1, 11.9, 1, 270, 360, 0, 1);
    Doc2D->ksLineSeg(25.1, 11.9, 25.1, 15.5, 1);
    Doc2D->ksLineSeg(25.1, 15.5, 27.1, 17.5, 1);
    Doc2D->ksLineSeg(27.1, 17.5, 237, 17.5, 1);
    Doc2D->ksLineSeg(237, 17.5, 237, 17, 1);
    Doc2D->ksLineSeg(237, 17, 245, 17, 1);
    Doc2D->ksLineSeg(245, 17, 245, 9, 1);
    Doc2D->ksLineSeg(245, 9, 252, 9, 1);
    Doc2D->ksLineSeg(252, 9, 252, 17.5, 1);
    Doc2D->ksLineSeg(252, 17.5, 254.5, 17.5, 1);
    Doc2D->ksLineSeg(254.5, 17.5, 260, 12, 1);
    Doc2D->ksLineSeg(260, 12, 277, 12, 1);
    Doc2D->ksLineSeg(277, 12, 280, 9, 1);
    Doc2D->ksLineSeg(280, 9, 280, 0, 1);

    sketchDef->EndEdit();

    ksEntityPtr BossRot = Part->NewEntity(Kompas6Constants3D::o3d_bossRotated);
    ksBossRotatedDefinitionPtr BossRotDef = BossRot->GetDefinition();
    BossRotDef->SetSideParam(TRUE, 360);
    BossRotDef->directionType* (Kompas6Constants3D::dtNormal);
    BossRotDef->SetSketch(ISketchEntity);
    BossRot->Create();

    ksEntityPtr Offset = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr OffsetDef = Offset->GetDefinition();
    OffsetDef->direction = true;
    OffsetDef->offset = -280;
    OffsetDef->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeYOZ));
    Offset->Create();

    ksEntityPtr ISketchEntity1 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
    sketchDef1->SetPlane(Offset);
    ISketchEntity1->Create();
    ksDocument2DPtr Doc2D1 = sketchDef1->BeginEdit();

    Doc2D1->ksLineSeg(-2.5, 9, 2.5, 9, 1);
    Doc2D1->ksLineSeg(2.5, 9, 2.5, 17.5, 1);
    Doc2D1->ksLineSeg(2.5, 17.5, -2.5, 17.5, 1);
    Doc2D1->ksLineSeg(-2.5, 17.5, -2.5, 9, 1);

    sketchDef1->EndEdit();

    ksEntityPtr CutExtr = Part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef = CutExtr->GetDefinition();
    CutExtrDef->SetSideParam(FALSE, 0, 25, 0, FALSE);
    CutExtrDef->SetSketch(ISketchEntity1);
    CutExtr->Create();

    kompass.Detach();

}

// Запуск программы: CTRL+F5 или меню "Отладка" > "Запуск без отладки"
// Отладка программы: F5 или меню "Отладка" > "Запустить отладку"

// Советы по началу работы 
//   1. В окне обозревателя решений можно добавлять файлы и управлять ими.
//   2. В окне Team Explorer можно подключиться к системе управления версиями.
//   3. В окне "Выходные данные" можно просматривать выходные данные сборки и другие сообщения.
//   4. В окне "Список ошибок" можно просматривать ошибки.
//   5. Последовательно выберите пункты меню "Проект" > "Добавить новый элемент", чтобы создать файлы кода, или "Проект" > "Добавить существующий элемент", чтобы добавить в проект существующие файлы кода.
//   6. Чтобы снова открыть этот проект позже, выберите пункты меню "Файл" > "Открыть" > "Проект" и выберите SLN-файл.
