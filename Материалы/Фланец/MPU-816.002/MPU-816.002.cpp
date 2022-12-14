//Программа построения детали MPU-816.002
//Автор: Гаврилова Е.С.
//Дата создания : 11.11.2022
//Дата последнего изменения : 11.11.2022

#include <iostream>
#import "C:\Program Files\ASCON\KOMPAS-3D v21 Study\Bin\\kAPI5.tlb"
#import "C:\Program Files\ASCON\KOMPAS-3D v21 Study\Bin\\ksConstants3D.tlb"

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
    //перый эскиз
    Doc2D->ksLineSeg(-(52.5 - 12.5), 52.5, 52.5 - 12.5, 52.5, 1);
    Doc2D->ksLineSeg(52.5, 52.5 - 12.5, 52.5, -(52.5 - 12.5), 1);
    Doc2D->ksLineSeg(52.5 - 12.5, -(52.5), -(52.5 - 12.5), -(52.5), 1);
    Doc2D->ksLineSeg(-(52.5), -(52.5 - 12.5), -(52.5), 52.5 - 12.5, 1);

    Doc2D->ksArcByAngle(52.5 - 12.5, 52.5 - 12.5, 12.5, 0, 90, 1, 1);
    Doc2D->ksArcByAngle(-(52.5 - 12.5), 52.5 - 12.5, 12.5, 90, 180, 1, 1);
    Doc2D->ksArcByAngle(52.5 - 12.5, -(52.5 - 12.5), 12.5, 270, 0, 1, 1);
    Doc2D->ksArcByAngle(-(52.5 - 12.5), -(52.5 - 12.5), 12.5, 180, 270, 1, 1);
    sketchDef->EndEdit();
    //выдавливание основного эскиза
    ksEntityPtr BossExtr = Part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef = BossExtr->GetDefinition();
    BossExtrDef->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 25, 0, FALSE);
    BossExtrDef->directionType* (Kompas6Constants3D::dtNormal);
    BossExtrDef->SetSketch(ISketchEntity);
    BossExtr->Create();

    //эскиз центрального отверистия
    ksEntityPtr ISketchEntity1 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
    sketchDef1->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
    ISketchEntity1->Create();
    ksDocument2DPtr Doc2D1 = sketchDef1->BeginEdit();

    Doc2D1->ksCircle(0, 0, 45, 1);

    sketchDef1->EndEdit();
    // выдавливание центрального отверстия
    ksEntityPtr CutExtr = Part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef = CutExtr->GetDefinition();
    CutExtrDef->cut* (TRUE);
    CutExtrDef->directionType* (Kompas6Constants3D::dtNormal);
    CutExtrDef->SetSideParam(FALSE, Kompas6Constants3D::etThroughAll, 20, 0, FALSE);
    CutExtrDef->SetSketch(ISketchEntity1);
    CutExtr->Create();
    /*
    круговые отверстия без фасок
    ksEntityPtr ISketchEntity2 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef2 = ISketchEntity2->GetDefinition();
    sketchDef2->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
    ISketchEntity2->Create();
    ksDocument2DPtr Doc2D2 = sketchDef2->BeginEdit();

    Doc2D2->ksCircle(40, 40, 6, 1);
    sketchDef2->EndEdit();

    ksEntityPtr CutExtr1 = Part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
    ksCutExtrusionDefinitionPtr CutExtrDef1 = CutExtr1->GetDefinition();
    CutExtrDef1->cut* (TRUE);
    CutExtrDef1->directionType* (Kompas6Constants3D::dtNormal);
    CutExtrDef1->SetSideParam(FALSE, Kompas6Constants3D::etThroughAll, 25, 0, FALSE);
    CutExtrDef1->SetSketch(ISketchEntity2);
    CutExtr1->Create();


    ksEntityPtr CircArr = Part->NewEntity(Kompas6Constants3D::o3d_circularCopy);
    ksCircularCopyDefinitionPtr CircDef = CircArr->GetDefinition();
    CircDef->SetAxis(Part->GetDefaultEntity(Kompas6Constants3D::o3d_axisOZ));
    CircDef->SetCopyParamAlongDir(4, 360, true, false);
    ksEntityCollectionPtr ColEn = CircDef->GetOperationArray();
    ColEn->Add(CutExtr1);
    CircArr->Create();
*/



//вспомогательная плоскость
    ksEntityPtr ISketchEntity0 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef0 = ISketchEntity0->GetDefinition();
    sketchDef0->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOZ));
    ISketchEntity0->Create();

    ksEntityPtr Offset1 = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr OffsetDef1 = Offset1->GetDefinition();
    OffsetDef1->direction = true;
    OffsetDef1->offset = 52.5 - 12.5;
    OffsetDef1->SetPlane(ISketchEntity0);
    Offset1->Create();

    ksEntityPtr ISketchEntity01 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef01 = ISketchEntity01->GetDefinition();
    sketchDef01->SetPlane(Offset1);
    ISketchEntity01->Create();
    ksDocument2DPtr Doc2D01 = sketchDef01->BeginEdit();

    //эскиз кругового отверстия
    Doc2D01->ksLineSeg(40, 0, 40, -25, 1);
    Doc2D01->ksLineSeg(40, -25, 40 + 6, -25, 1);
    Doc2D01->ksLineSeg(40 + 6, -25, 40 + 6 - 1.6, -(25 - 1.6), 1);
    Doc2D01->ksLineSeg(40 + 6 - 1.6, -(25 - 1.6), 40 + 6 - 1.6, 0, 1);
    Doc2D01->ksLineSeg(40 + 6 - 1.6, 0, 40, 0, 1);
    Doc2D01->ksLineSeg(40, 0, 40, -25, 3);
    sketchDef01->EndEdit();

    // вырезание кругового отверстия
    ksEntityPtr CutRExtr1 = Part->NewEntity(Kompas6Constants3D::o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRExtrDef1 = CutRExtr1->GetDefinition();
    CutRExtrDef1->cut* (TRUE);
    CutRExtrDef1->directionType* (Kompas6Constants3D::dtNormal);
    CutRExtrDef1->SetSideParam(TRUE, 360);
    CutRExtrDef1->SetSketch(ISketchEntity01);
    CutRExtr1->Create();

    //массив отверстий
    ksEntityPtr CircArr = Part->NewEntity(Kompas6Constants3D::o3d_circularCopy);
    ksCircularCopyDefinitionPtr CircDef = CircArr->GetDefinition();
    CircDef->SetAxis(Part->GetDefaultEntity(Kompas6Constants3D::o3d_axisOZ));
    CircDef->SetCopyParamAlongDir(4, 360, true, false);
    ksEntityCollectionPtr ColEn = CircDef->GetOperationArray();
    ColEn->Add(CutRExtr1);
    CircArr->Create();

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
