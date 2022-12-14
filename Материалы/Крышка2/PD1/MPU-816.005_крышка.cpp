// 005.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//

#include <iostream>
#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\kAPI5.tlb"
#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\ksConstants3D.tlb"

using namespace std;
using namespace Kompas6API5;


int main()
{
	CoInitialize(NULL);
	KompasObjectPtr kompass;
	HRESULT hRes;
	hRes = kompass.GetActiveObject(L"Kompas.Application.5");
	if (FAILED(hRes))
		kompass.CreateInstance(L"Kompas.Application.5");
	kompass->Visible = true;
	ksDocument3DPtr idoc3D = kompass->Document3D();
	idoc3D->Create(false, true);
	idoc3D = kompass->ActiveDocument3D();

	ksPartPtr part;
	part = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);

	//эскиз 1
	ksEntityPtr ISketchEntity = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef = ISketchEntity->GetDefinition();
	sketchDef->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity->Create();
	ksDocument2DPtr doc2d = sketchDef->BeginEdit();
	doc2d->ksLineSeg(-52.5, -40, -52.5, 40, 1);//левая прямая


	doc2d->ksLineSeg(-40, 52.5, 40, 52.5, 1);//верхняя прямая

	doc2d->ksLineSeg(52.5, 40, 52.5, -40, 1);//правая прямая

	doc2d->ksLineSeg(40, -52.5, -40, -52.5, 1);//нижняя прямая
	doc2d->ksArcByAngle(52.5 - 12.5, 52.5 - 12.5, 12.5, 0, 90, 1, 1);
	doc2d->ksArcByAngle(-(52.5 - 12.5), 52.5 - 12.5, 12.5, 90, 180, 1, 1);
	doc2d->ksArcByAngle(52.5 - 12.5, -(52.5 - 12.5), 12.5, 270, 0, 1, 1);
	doc2d->ksArcByAngle(-(52.5 - 12.5), -(52.5 - 12.5), 12.5, 180, 270, 1, 1);
	sketchDef->EndEdit();

	//Выдавливание 1 основного эскиза
	ksEntityPtr BossExtr = part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion);
	ksBossExtrusionDefinitionPtr BossExtrDef = BossExtr->GetDefinition();
	BossExtrDef->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 25, 0, FALSE);
	BossExtrDef->directionType* (Kompas6Constants3D::dtNormal);
	BossExtrDef->SetSketch(ISketchEntity);
	BossExtr->Create();


	//Эскиз для основного отверстия
	ksEntityPtr ISketchEntity1 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
	sketchDef1->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity1->Create();
	ksDocument2DPtr doc2d1 = sketchDef1->BeginEdit();

	doc2d1->ksCircle(0, 0, 35 / 2, 1);

	sketchDef1->EndEdit();
	//вырез основного отверстия
	ksEntityPtr CutExtr1 = part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr CutExtrDef1 = CutExtr1->GetDefinition();
	CutExtrDef1->SetSideParam(FALSE, Kompas6Constants3D::etThroughAll, 25, 0, FALSE);
	CutExtrDef1->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef1->SetSketch(ISketchEntity1);
	CutExtr1->Create();


	// эскиз  для выдавливания сверху крышки
	ksEntityPtr ISketchEntity2 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef2 = ISketchEntity2->GetDefinition();
	sketchDef2->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity2->Create();
	ksDocument2DPtr doc2d2 = sketchDef2->BeginEdit();

	doc2d2->ksCircle(0, 0, 49 / 2, 1);
	sketchDef2->EndEdit();

	//выдавливание вниз на 7 на крышку
	ksEntityPtr CutExtr2 = part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr CutExtrDef2 = CutExtr2->GetDefinition();
	CutExtrDef2->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 7, 0, FALSE);
	CutExtrDef2->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef2->SetSketch(ISketchEntity2);
	CutExtr2->Create();

	// эскиз  для выдавливания вверх на крышке
	ksEntityPtr ISketchEntity3 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef3 = ISketchEntity3->GetDefinition();
	sketchDef3->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity3->Create();
	ksDocument2DPtr doc2d3 = sketchDef3->BeginEdit();

	doc2d3->ksCircle(0, 0, 60 / 2, 1);
	doc2d3->ksCircle(0, 0, 35 / 2, 1);
	sketchDef3->EndEdit();

	ksEntityPtr BossExtr2 = part->NewEntity(Kompas6Constants3D::o3d_baseExtrusion);
	ksBaseExtrusionDefinitionPtr BossExtrDef2 = BossExtr2->GetDefinition();
	BossExtrDef2->SetSideParam(FALSE, Kompas6Constants3D::etBlind, 5, 0, FALSE);
	BossExtrDef2->directionType = Kompas6Constants3D::dtReverse;
	BossExtrDef2->SetSketch(ISketchEntity3);
	BossExtr2->Create();




	//вспомогательная плоскость для выдавливания внизу
	ksEntityPtr ISketchEntity0 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef0 = ISketchEntity0->GetDefinition();
	sketchDef0->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity0->Create();

	ksEntityPtr Offset1 = part->NewEntity(Kompas6Constants3D::o3d_planeOffset);
	ksPlaneOffsetDefinitionPtr OffsetDef1 = Offset1->GetDefinition();
	OffsetDef1->direction = true;
	OffsetDef1->offset = 25;
	OffsetDef1->SetPlane(ISketchEntity0);
	Offset1->Create();

	ksEntityPtr ISketchEntity01 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef01 = ISketchEntity01->GetDefinition();
	sketchDef01->SetPlane(Offset1);
	ISketchEntity01->Create();
	ksDocument2DPtr doc2d01 = sketchDef01->BeginEdit();

	doc2d01->ksCircle(0, 0, 60 / 2, 1);
	doc2d01->ksCircle(0, 0, 80 / 2, 1);
	sketchDef01->EndEdit();

	//выдавливание снизу
	ksEntityPtr BossExtr3 = part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion);
	ksBossExtrusionDefinitionPtr BossExtrDef3 = BossExtr3->GetDefinition();
	BossExtrDef3->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 15, 0, FALSE);
	BossExtrDef3->directionType* (Kompas6Constants3D::dtNormal);
	BossExtrDef3->SetSketch(ISketchEntity01);
	BossExtr3->Create();

	//эскиз для выреза сбоку под резьбу
	ksEntityPtr ISketchEntity4 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef4 = ISketchEntity4->GetDefinition();
	sketchDef4->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOZ));
	ISketchEntity4->Create();
	ksDocument2DPtr doc2d4 = sketchDef4->BeginEdit();

	doc2d4->ksLineSeg(-52.5, -12.5, -20.5, -12.5, 3);
	doc2d4->ksLineSeg(-52.5, -12.5, -52.5, -21.028, 1);
	doc2d4->ksLineSeg(-52.5, -21.028, -36.5, -19.899, 1);
	doc2d4->ksLineSeg(-36.5, -19.899, -36.5, -16.5, 1);
	doc2d4->ksLineSeg(-36.5, -16.5, -22.809, -16.5, 1);
	doc2d4->ksLineSeg(-22.809, -16.5, -20.5, -12.5, 1);
	sketchDef4->EndEdit();


	//вырез вращением под резьбу
	ksEntityPtr CutRot = part->NewEntity(Kompas6Constants3D::o3d_cutRotated);
	ksCutRotatedDefinitionPtr CutRotDef = CutRot->GetDefinition();
	CutRotDef->cut = TRUE;
	CutRotDef->SetSideParam(TRUE, 360);
	CutRotDef->directionType* (Kompas6Constants3D::dtNormal);
	CutRotDef->SetSketch(ISketchEntity4);
	CutRot->Create();


	//эскиз выреза отверстия внизу в резьбу
	ksEntityPtr ISketchEntity5 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef5 = ISketchEntity5->GetDefinition();
	sketchDef5->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOZ));
	ISketchEntity5->Create();
	ksDocument2DPtr doc2d5 = sketchDef5->BeginEdit();

	doc2d5->ksLineSeg(-26, 0 - 25, -26, 12.5 - 25, 3);
	doc2d5->ksLineSeg(-26, 12.5 - 25, -30, 8.5 - 25, 1);
	doc2d5->ksLineSeg(-30, 8.5 - 25, -30, 0 - 25, 1);
	doc2d5->ksLineSeg(-30, 0 - 25, -26, 0 - 25, 1);

	sketchDef5->EndEdit();

	//вырез ращением отвестия внизу под резьбу
	ksEntityPtr CutRot1 = part->NewEntity(Kompas6Constants3D::o3d_cutRotated);
	ksCutRotatedDefinitionPtr CutRotDef1 = CutRot1->GetDefinition();
	CutRotDef1->cut = TRUE;
	CutRotDef1->SetSideParam(TRUE, 360);
	CutRotDef1->directionType* (Kompas6Constants3D::dtNormal);
	CutRotDef1->SetSketch(ISketchEntity5);
	CutRot1->Create();

	//эскиз для отверстий сверху крышки(верхняя часть цековки)
	ksEntityPtr ISketchEntity6 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef6 = ISketchEntity6->GetDefinition();
	sketchDef6->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity6->Create();
	ksDocument2DPtr doc2d6 = sketchDef6->BeginEdit();

	doc2d6->ksCircle(40, 40, 10, 1);

	sketchDef6->EndEdit();

	ksEntityPtr CutExtr3 = part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr CutExtrDef3 = CutExtr3->GetDefinition();
	CutExtrDef3->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 16, 0, FALSE);
	CutExtrDef3->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef3->SetSketch(ISketchEntity6);
	CutExtr3->Create();

	//нижняя часть цековки
	ksEntityPtr ISketchEntity7 = part->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef7 = ISketchEntity7->GetDefinition();
	sketchDef7->SetPlane(part->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity7->Create();
	ksDocument2DPtr doc2d7 = sketchDef7->BeginEdit();

	doc2d7->ksCircle(40, 40, 13 / 2, 1);

	sketchDef7->EndEdit();

	//почему-то не выдавливает нормально
	ksEntityPtr CutExtr4 = part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr CutExtrDef4 = CutExtr4->GetDefinition();
	CutExtrDef4->SetSideParam(false, Kompas6Constants3D::etThroughAll, 25, 0, FALSE);
	CutExtrDef4->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef4->SetSketch(ISketchEntity7);
	CutExtr4->Create();


	//ksEntityPtr EnAxis = part->NewEntity(Kompas6Constants3D::o3d_axisOX);
	//EnAxis->Create();

	// массив центовок, тоже почему-то не работает
	ksEntityPtr MeshCopy = part->NewEntity(Kompas6Constants3D::o3d_meshCopy);
	ksMeshCopyDefinitionPtr MeshCopyDef = MeshCopy->GetDefinition();
	//MeshCopyDef->insideFlag = FALSE;
	MeshCopyDef->SetCopyParamAlongAxis(TRUE, -90, 2, 80, FALSE);
	//MeshCopyDef->SetAxis2(EnAxis);
	MeshCopyDef->SetCopyParamAlongAxis(FALSE, -90, 2, 80, FALSE);
	ksEntityCollectionPtr EnColPartDef = MeshCopyDef->OperationArray();
	//EnColPartDef->Clear();
	EnColPartDef->Add(CutExtr4);
	EnColPartDef->Add(CutExtr3);

	MeshCopy->Create();

	//скругление снизу
	ksEntityPtr Fillet = part->NewEntity(Kompas6Constants3D::o3d_fillet);
	ksFilletDefinitionPtr FilletDef = Fillet->GetDefinition();
	FilletDef->radius = 3;
	FilletDef->tangent = FALSE;

	ksEntityCollectionPtr EnColPart = part->EntityCollection(Kompas6Constants3D::o3d_edge);
	ksEntityCollectionPtr EnColFillet = FilletDef->array();
	EnColFillet->Clear();

	EnColPart->SelectByPoint(0, 40, 40);
	EnColFillet->Add(EnColPart->GetByIndex(0));

	Fillet->Create();

	//скругление сверху 1
	ksEntityPtr Fillet1 = part->NewEntity(Kompas6Constants3D::o3d_fillet);
	ksFilletDefinitionPtr FilletDef1 = Fillet1->GetDefinition();
	FilletDef1->radius = 2.5;
	FilletDef1->tangent = FALSE;

	ksEntityCollectionPtr EnColPart1 = part->EntityCollection(Kompas6Constants3D::o3d_edge);
	ksEntityCollectionPtr EnColFillet1 = FilletDef1->array();
	EnColFillet1->Clear();

	EnColPart1->SelectByPoint(0, 30, -5);
	EnColFillet1->Add(EnColPart1->GetByIndex(0));

	Fillet1->Create();

	//скругление сверху 2
	ksEntityPtr Fillet2 = part->NewEntity(Kompas6Constants3D::o3d_fillet);
	ksFilletDefinitionPtr FilletDef2 = Fillet2->GetDefinition();
	FilletDef2->radius = 2.5;
	FilletDef2->tangent = FALSE;

	ksEntityCollectionPtr EnColPart2 = part->EntityCollection(Kompas6Constants3D::o3d_edge);
	ksEntityCollectionPtr EnColFillet2 = FilletDef2->array();
	EnColFillet2->Clear();

	EnColPart2->SelectByPoint(0, 30, 0);
	EnColFillet2->Add(EnColPart2->GetByIndex(0));

	Fillet2->Create();

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
