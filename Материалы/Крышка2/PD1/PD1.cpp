#include <iostream>
#include <atlbase.h>
#include <OleAuto.h>

using namespace std;

#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\kAPI5.tlb"
#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\ksConstants3D.tlb"
#import "C:\\Program Files\\ASCON\\KOMPAS-3D v21 Study\\Bin\\ksConstants.tlb"

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

	ksPartPtr part;// эскиз 1
	part = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
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
	ksPartPtr part1;
	part1 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	ksEntityPtr ISketchEntity1 = part1->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef1 = ISketchEntity1->GetDefinition();
	sketchDef1->SetPlane(part1->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity1->Create();
	ksDocument2DPtr doc2d1 = sketchDef1->BeginEdit();

	doc2d1->ksCircle(0, 0, 35/2, 1);

	sketchDef1->EndEdit();
	//вырез основного отверстия
	ksEntityPtr CutExtr1 = part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr CutExtrDef1 = CutExtr1->GetDefinition();
	CutExtrDef1->SetSideParam(FALSE, Kompas6Constants3D::etThroughAll, 25, 0, FALSE);
	CutExtrDef1->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef1->SetSketch(ISketchEntity1);
	CutExtr1->Create();


	// эскиз 2 для выдавливания сверху крышки
	ksPartPtr part2;
	part2 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	ksEntityPtr ISketchEntity2 = part2->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef2 = ISketchEntity2->GetDefinition();
	sketchDef2->SetPlane(part2->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity2->Create();
	ksDocument2DPtr doc2d2 = sketchDef2->BeginEdit();

	doc2d2->ksCircle(0, 0, 49/2, 1);
	sketchDef2->EndEdit();

	//выдавливание2 вниз на 7 на крышку
	ksEntityPtr CutExtr2 = part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr CutExtrDef2 = CutExtr2->GetDefinition();
	CutExtrDef2->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 7, 0, FALSE);
	CutExtrDef2->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef2->SetSketch(ISketchEntity2);
	CutExtr2->Create();

	//// эскиз 3 для выдавливания 
	//ksPartPtr part3;
	//part3 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	//ksEntityPtr ISketchEntity3 = part3->NewEntity(Kompas6Constants3D::o3d_sketch);
	//ksSketchDefinitionPtr sketchDef3 = ISketchEntity3->GetDefinition();
	//sketchDef3->SetPlane(part3->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	//ISketchEntity3->Create();
	//ksDocument2DPtr doc2d3 = sketchDef3->BeginEdit();

	//doc2d3->ksCircle(0, 0, 60/2, 1);
	//doc2d3->ksCircle(0, 0, 35 / 2, 1);
	//sketchDef3->EndEdit();

	//ksEntityPtr BossExtr2 = part->NewEntity(Kompas6Constants3D::o3d_bossExtrusion);
	//ksBossExtrusionDefinitionPtr BossExtrDef2 = BossExtr2->GetDefinition();
	//BossExtrDef2->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 5, 0, FALSE);
	//BossExtrDef2->directionType* (Kompas6Constants3D::dtReverse);
	//BossExtrDef2->SetSketch(ISketchEntity3);
	//BossExtr2->Create();




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

	doc2d01->ksCircle(0, 0, 60/2, 1);
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
	ksPartPtr part4;
	part4 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	ksEntityPtr ISketchEntity4 = part4->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef4 = ISketchEntity4->GetDefinition();
	sketchDef4->SetPlane(part4->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOZ));
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
	CutRotDef->SetSideParam(TRUE, 360);
	CutRotDef->directionType* (Kompas6Constants3D::dtNormal);
	CutRotDef->SetSketch(ISketchEntity4);
	CutRot->Create();


	//эскиз выреза отверстия внизу в резьбу
	ksPartPtr part5;
	part5 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	ksEntityPtr ISketchEntity5 = part5->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef5 = ISketchEntity5->GetDefinition();
	sketchDef5->SetPlane(part5->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOZ));
	ISketchEntity5->Create();
	ksDocument2DPtr doc2d5 = sketchDef5->BeginEdit();

	doc2d5->ksLineSeg(-26, 0-25, -26, 12.5-25, 3);
	doc2d5->ksLineSeg(-26, 12.5-25, -30, 8.5-25, 1);
	doc2d5->ksLineSeg(-30, 8.5-25, -30, 0-25, 1);
	doc2d5->ksLineSeg(-30, 0-25, -26, 0-25, 1);

	sketchDef5->EndEdit();

	//вырез ращением отвестия внизу под резьбу
	ksEntityPtr CutRot1 = part->NewEntity(Kompas6Constants3D::o3d_cutRotated);
	ksCutRotatedDefinitionPtr CutRotDef1 = CutRot1->GetDefinition();
	CutRotDef1->SetSideParam(TRUE, 360);
	CutRotDef1->directionType* (Kompas6Constants3D::dtNormal);
	CutRotDef1->SetSketch(ISketchEntity5);
	CutRot1->Create();

	//эскиз для отверстий сверху крышки
	ksPartPtr part6;
	part6 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	ksEntityPtr ISketchEntity6 = part6->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef6 = ISketchEntity6->GetDefinition();
	sketchDef6->SetPlane(part6->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
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
	ksPartPtr part7;
	part7 = idoc3D->GetPart(Kompas6Constants3D::pTop_Part);
	ksEntityPtr ISketchEntity7 = part7->NewEntity(Kompas6Constants3D::o3d_sketch);
	ksSketchDefinitionPtr sketchDef7 = ISketchEntity7->GetDefinition();
	sketchDef7->SetPlane(part7->GetDefaultEntity(Kompas6Constants3D::o3d_planeXOY));
	ISketchEntity7->Create();
	ksDocument2DPtr doc2d7 = sketchDef7->BeginEdit();

	doc2d7->ksCircle(40, 40, 13/2, 1);

	sketchDef7->EndEdit();

	ksEntityPtr CutExtr4 = part->NewEntity(Kompas6Constants3D::o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr CutExtrDef4 = CutExtr4->GetDefinition();
	CutExtrDef4->SetSideParam(TRUE, Kompas6Constants3D::etBlind, 25, 0, FALSE);
	CutExtrDef4->directionType* (Kompas6Constants3D::dtNormal);
	CutExtrDef4->SetSketch(ISketchEntity7);
	CutExtr4->Create();
	kompass.Detach();
}

