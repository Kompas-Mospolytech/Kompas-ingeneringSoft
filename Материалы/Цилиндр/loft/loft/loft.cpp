// loft.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
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

    Doc2D->ksCircle(0, 0, 30, 1);

    sketchDef->EndEdit();

    ksEntityPtr Offset = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr OffsetDef = Offset->GetDefinition();
    OffsetDef->direction = true;
    OffsetDef->offset = 20;
    OffsetDef->SetPlane(ISketchEntity);
    Offset->Create();


    ksEntityPtr ISketchEntity1 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
    sketchDef1->SetPlane(Offset);
    ISketchEntity1->Create();
    ksDocument2DPtr Doc2D1 = sketchDef1->BeginEdit();

    Doc2D1->ksCircle(0, 0, 20, 1);

    sketchDef1->EndEdit();

    ksEntityPtr Offset1 = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr OffsetDef1 = Offset1->GetDefinition();
    OffsetDef1->direction = true;
    OffsetDef1->offset = 40;
    OffsetDef1->SetPlane(ISketchEntity);
    Offset1->Create();

    ksEntityPtr ISketchEntity2 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef2 = ISketchEntity2->GetDefinition();
    sketchDef2->SetPlane(Offset1);
    ISketchEntity2->Create();
    ksDocument2DPtr Doc2D2 = sketchDef2->BeginEdit();

    Doc2D2->ksCircle(0, 0, 25, 1);

    sketchDef2->EndEdit();

    ksEntityPtr BossLoft = Part->NewEntity(Kompas6Constants3D::o3d_bossLoft);
    ksBossLoftDefinitionPtr BossLoftDef = BossLoft->GetDefinition();
    BossLoftDef->SetLoftParam(FALSE, TRUE, TRUE);
    ksEntityCollectionPtr Colen = BossLoftDef->Sketchs();
    Colen->Clear();

    Colen->Add(ISketchEntity);
    Colen->Add(ISketchEntity1);
    Colen->Add(ISketchEntity2);

    BossLoft->Create();


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
