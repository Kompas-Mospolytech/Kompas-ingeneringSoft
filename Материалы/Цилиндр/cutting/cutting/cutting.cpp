// cutting.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//

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
    sketchDef->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
    ISketchEntity->Create();
    ksDocument2DPtr Doc2D = sketchDef->BeginEdit();


    Kompas6API5::ksMathematic2DPtr Mathematic2D;
    Mathematic2D = *new Kompas6API5::ksMathematic2DPtr(kompass->GetMathematic2D());

    Doc2D->ksLineSeg(0, 10, 5, 10, 1);
    Doc2D->ksLineSeg(5, 10, 5, 15, 1);
    Doc2D->ksLineSeg(5, 15, 0, 15, 1);
    Doc2D->ksLineSeg(0, 15, 0, 10, 1);

    sketchDef->EndEdit();

    ksEntityPtr BossExtr = Part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef = BossExtr->GetDefinition();
    BossExtrDef->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 20, 0, FALSE);
    BossExtrDef->directionType* (Kompas6Constants3D::dtNormal);
    BossExtrDef->SetSketch(ISketchEntity);
    BossExtr->Create();

    ksEntityPtr ISketchEntity1 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
    sketchDef1->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
    ISketchEntity1->Create();
    ksDocument2DPtr Doc2D1 = sketchDef1->BeginEdit();

    Doc2D1->ksLineSeg(1, 11, 4, 11, 1);
    Doc2D1->ksLineSeg(4, 11, 4, 14, 1);
    Doc2D1->ksLineSeg(4, 14, 1, 14, 1);
    Doc2D1->ksLineSeg(1, 14, 1, 11, 1);

    sketchDef1->EndEdit();

    ksEntityPtr CutExtr = Part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef = CutExtr->GetDefinition();
    CutExtrDef->SetSideParam(FALSE, 1, 20, 0, FALSE);
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
