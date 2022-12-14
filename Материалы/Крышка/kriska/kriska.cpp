// kriska.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//

#include <iostream>
#import "E:\\kompas\\Bin\\kAPI5.tlb"
#import "E:\\kompas\\Bin\\ksConstants3D.tlb"
#import "E:\\kompas\\Bin\\ksConstants.tlb"
using namespace std;
using namespace Kompas6API5;

int main()
{
    CoInitialize(NULL);
    Kompas6API5::KompasObjectPtr kompass;
    HRESULT hRes = kompass.GetActiveObject(L"Kompas.Application.5");
    if (FAILED(hRes))
        kompass.CreateInstance(L"Kompas.Application.5");
    kompass->Visible = true;
    ksDocument3DPtr iDoc3d = kompass->Document3D();
    iDoc3d->Create(false, true);
    iDoc3d = kompass->ActiveDocument3D();
    ksPartPtr Part;
    Part = iDoc3d->GetPart(Kompas6Constants3D::pTop_Part);

    //эскиз1
    ksEntityPtr ISketchEntity = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef = ISketchEntity->GetDefinition();
    sketchDef->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
    ISketchEntity->Create();
    ksDocument2DPtr doc2d = sketchDef->BeginEdit();

    doc2d->ksArcByPoint(12.5, 12.5, 12.5, 0, 12.5, 12.5, 0, 2, 1);
    doc2d->ksLineSeg(12.5, 0, 92.5, 0, 1);
    doc2d->ksArcByPoint(92.5, 12.5, 12.5, 92.5, 0, 105, 12.5, 1, 1);
    doc2d->ksLineSeg(105, 12.5, 105, 92.5, 1);
    doc2d->ksArcByPoint(92.5, 92.5, 12.5, 105, 92.5, 92.5, 105, 1, 1);
    doc2d->ksLineSeg(92.5, 105, 12.5, 105, 1);
    doc2d->ksArcByPoint(12.5, 92.5, 12.5, 12.5, 105, 0, 92.5, 1, 1);
    doc2d->ksLineSeg(0, 92.5, 0, 12.5, 1);

    sketchDef->EndEdit();

    //выдавливание
    ksEntityPtr BossExtr = Part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef = BossExtr->GetDefinition();
    BossExtrDef->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 26, 0, FALSE);
    BossExtrDef->directionType* (Kompas6Constants3D::dtNormal);
    BossExtrDef->SetSketch(ISketchEntity);
    BossExtr->Create();
 
    //смещенная плоскость
    ksEntityPtr EnPlaneOffset = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef = EnPlaneOffset->GetDefinition();
    PlaneOffsetDef->direction = FALSE;
    PlaneOffsetDef->offset = 12.5;
    PlaneOffsetDef->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeYOZ));
    EnPlaneOffset->Create();

    //эскиз2
    ksEntityPtr ISketchEntity1 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
    sketchDef1->SetPlane(EnPlaneOffset);
    ISketchEntity1->Create();
    ksDocument2DPtr doc2d1 = sketchDef1->BeginEdit();

    doc2d1->ksLineSeg(-26, -2.5, -26, -12.5, 1);
    doc2d1->ksLineSeg(-26, -12.5, 0, -12.5, 3);
    doc2d1->ksLineSeg(0, -12.5, 0, -6, 1);
    doc2d1->ksLineSeg(0, -6, -10, -6, 1);
    doc2d1->ksLineSeg(-10, -6, -10, -2.5, 1);
    doc2d1->ksLineSeg(-10, -2.5, -26, -2.5, 1);

    sketchDef1->EndEdit();

    //вырезание вращением
    ksEntityPtr EnCutRot = Part->NewEntity(Kompas6Constants3D::o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef = EnCutRot->GetDefinition();
    CutRotDef->directionType* (Kompas6Constants3D::dtNormal);
    CutRotDef->cut = TRUE;
    CutRotDef->SetSideParam(TRUE, 360);
    CutRotDef->SetSketch(ISketchEntity1);
    EnCutRot->Create();

    ksEntityPtr EnAxis = Part->NewEntity(Kompas6Constants3D::o3d_axisOX);
    EnAxis->Create();
    //массив
    ksEntityPtr MeshCopy = Part->NewEntity(Kompas6Constants3D::o3d_meshCopy);
    ksMeshCopyDefinitionPtr MeshCopyDef = MeshCopy->GetDefinition();
    MeshCopyDef->insideFlag = FALSE;
    MeshCopyDef->SetCopyParamAlongAxis(TRUE, 0, 2, 80, FALSE);
    MeshCopyDef->SetAxis2(EnAxis);
    MeshCopyDef->SetCopyParamAlongAxis(FALSE, 90, 2, 80, FALSE);
    ksEntityCollectionPtr EnColPartDef = MeshCopyDef->OperationArray();
    EnColPartDef->Clear();
    EnColPartDef->Add(EnCutRot);
    MeshCopy->Create();

    //эскиз3
    ksEntityPtr ISketchEntity3 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef3 = ISketchEntity3->GetDefinition();
    sketchDef3->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
    ISketchEntity3->Create();
    ksDocument2DPtr doc2d3 = sketchDef3->BeginEdit();
    doc2d3->ksCircle(52.5, 52.5, 30, 1);
    doc2d3->ksCircle(52.5, 52.5, 40, 1);
    sketchDef3->EndEdit();

    //выдавливание
    ksEntityPtr BossExtr3 = Part->NewEntity(Kompas6Constants3D::o3d_baseExtrusion);
    ksBaseExtrusionDefinitionPtr BossExtrDef3 = BossExtr3->GetDefinition();
    BossExtrDef3->SetSideParam(FALSE, Kompas6Constants3D::etBlind, 15, 0, FALSE);
    BossExtrDef3->directionType=Kompas6Constants3D::dtReverse;
    BossExtrDef3->SetSketch(ISketchEntity3);
    BossExtr3->Create();

    //смещенная плоскость
    ksEntityPtr EnPlaneOffset1 = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef1 = EnPlaneOffset1->GetDefinition();
    PlaneOffsetDef1->direction = FALSE;
    PlaneOffsetDef1->offset = 62.5;
    PlaneOffsetDef1->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeYOZ));
    EnPlaneOffset1->Create();

    //эскиз4
    ksEntityPtr ISketchEntity4 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef4 = ISketchEntity4->GetDefinition();
    sketchDef4->SetPlane(EnPlaneOffset1);
    ISketchEntity4->Create();
    ksDocument2DPtr doc2d4 = sketchDef4->BeginEdit();

    doc2d4->ksLineSeg(-26, -18.5, -31, -18.5, 1);
    doc2d4->ksLineSeg(-31, -18.5, -68, -52.5, 1);
    //doc2d4->ksLineSeg(-31, -18.5, -60, -38, 1);
    //doc2d4->ksArcByPoint(-50, -52.5, 18, -60, -38, -68, -52.5, 1, 1);
    doc2d4->ksLineSeg(-68, -52.5, -58, -52.5, 1);
    doc2d4->ksArcBy3Points(-58, -52.5, -50, -44.5, -42, -52.5, 1);
    doc2d4->ksLineSeg(-42, -52.5, -26, -52.5, 1);
    doc2d4->ksLineSeg(-26, -52.5, -26, -18.5, 1);

    sketchDef4->EndEdit();

    //выдавливание
    ksEntityPtr BossExtr4 = Part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion);
    ksBossExtrusionDefinitionPtr BossExtrDef4 = BossExtr4->GetDefinition();
    BossExtrDef4->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 20, 0, FALSE);
    BossExtrDef4->directionType* (Kompas6Constants3D::dtNormal);
    BossExtrDef4->SetSketch(ISketchEntity4);
    BossExtr4->Create();

    //смещенная плоскость
    ksEntityPtr EnPlaneOffset2 = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef2 = EnPlaneOffset2->GetDefinition();
    PlaneOffsetDef2->direction = TRUE;
    PlaneOffsetDef2->offset = 52.5;
    PlaneOffsetDef2->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOZ));
    EnPlaneOffset2->Create();

    //зеркальное отражение
    ksEntityPtr EnMirrorOp = Part->NewEntity(Kompas6Constants3D::o3d_mirrorOperation);
    ksMirrorCopyDefinitionPtr MirrorCopyDef = EnMirrorOp->GetDefinition();
    MirrorCopyDef->SetPlane(EnPlaneOffset2);
    ksEntityCollectionPtr EnColl = MirrorCopyDef->GetOperationArray();
    EnColl->Clear();
    EnColl->Add(BossExtr4);
    EnMirrorOp->Create();

    //смещенная плоскость
    ksEntityPtr EnPlaneOffset5 = Part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
    ksPlaneOffsetDefinitionPtr PlaneOffsetDef5 = EnPlaneOffset5->GetDefinition();
    PlaneOffsetDef5->direction = FALSE;
    PlaneOffsetDef5->offset = 52.5;
    PlaneOffsetDef5->SetPlane(Part->GetDefaultEntity(Kompas6Constants3D::o3d_planeYOZ));
    EnPlaneOffset5->Create();

    //эскиз5
    ksEntityPtr ISketchEntity5 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef5 = ISketchEntity5->GetDefinition();
    sketchDef5->SetPlane(EnPlaneOffset5);
    ISketchEntity5->Create();
    ksDocument2DPtr doc2d5 = sketchDef5->BeginEdit();

    doc2d5->ksLineSeg(-13, 0, -21.53, 0, 1);
    doc2d5->ksLineSeg(-21.53, 0, -20.04, -16, 1);
    doc2d5->ksLineSeg(-20.04, -16, -17, -16, 1);
    doc2d5->ksLineSeg(-17, -16, -17, -32, 1);
    doc2d5->ksLineSeg(-17, -32, -13, -34.31, 1);
    doc2d5->ksLineSeg(-13, -34.31, -13, 0, 3);
    sketchDef5->EndEdit();

    //вырезание вращением
    ksEntityPtr EnCutRot5 = Part->NewEntity(Kompas6Constants3D::o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef5 = EnCutRot5->GetDefinition();
    CutRotDef5->directionType* (Kompas6Constants3D::dtNormal);
    CutRotDef5->cut = TRUE;
    CutRotDef5->SetSideParam(TRUE, 360);
    CutRotDef5->SetSketch(ISketchEntity5);
    EnCutRot5->Create();

    //эскиз6
    ksEntityPtr ISketchEntity6 = Part->NewEntity(Kompas6Constants3D::o3d_sketch);
    ksSketchDefinitionPtr sketchDef6 = ISketchEntity6->GetDefinition();
    sketchDef6->SetPlane(EnPlaneOffset5);
    ISketchEntity6->Create();
    ksDocument2DPtr doc2d6 = sketchDef6->BeginEdit();

    doc2d6->ksLineSeg(0, -22.5, -10, -22.5, 1);
    doc2d6->ksLineSeg(-10, -22.5, -10, -26.5, 1);
    doc2d6->ksLineSeg(-10, -26.5, 0, -26.5, 3);
    doc2d6->ksLineSeg(0, -26.5, 0, -22.5, 1);
    sketchDef6->EndEdit();

    //вырезание вращением
    ksEntityPtr EnCutRot6 = Part->NewEntity(Kompas6Constants3D::o3d_cutRotated);
    ksCutRotatedDefinitionPtr CutRotDef6 = EnCutRot6->GetDefinition();
    CutRotDef6->directionType* (Kompas6Constants3D::dtNormal);
    CutRotDef6->cut = TRUE;
    CutRotDef6->SetSideParam(TRUE, 360);
    CutRotDef6->SetSketch(ISketchEntity6);
    EnCutRot6->Create();

    //скругление1
    ksEntityPtr Fillet = Part->NewEntity(Kompas6Constants3D::o3d_fillet);
    ksFilletDefinitionPtr FilletDef = Fillet->GetDefinition();
    FilletDef->radius = 3;
    FilletDef->tangent = FALSE;

    ksEntityCollectionPtr EnColPart = Part->EntityCollection(Kompas6Constants3D::o3d_edge);
    ksEntityCollectionPtr EnColFillet = FilletDef->array();
    EnColFillet->Clear();
    
    EnColPart->SelectByPoint(52.5, 82.5, 0);
    EnColFillet->Add(EnColPart->GetByIndex(0));
    Fillet->Create();

    //скругление2
    ksEntityPtr Fillet2 = Part->NewEntity(Kompas6Constants3D::o3d_fillet);
    ksFilletDefinitionPtr FilletDef2 = Fillet2->GetDefinition();
    FilletDef2->radius = 8;
    FilletDef2->tangent = FALSE;

    ksEntityCollectionPtr EnColPart2 = Part->EntityCollection(Kompas6Constants3D::o3d_edge);
    ksEntityCollectionPtr EnColFillet2 = FilletDef2->array();
    EnColFillet2->Clear();

    EnColPart2->SelectByPoint(52.5, 52.5, 68);
    EnColFillet2->Add(EnColPart2->GetByIndex(0));

    Fillet2->Create();
    
    cout << "Hello World!\n";
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
