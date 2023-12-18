using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Kompas6Constants;
using Kompas6Constants3D;
using Kompas6API5;
using System.Runtime.CompilerServices;
using KompasAPI7;
using CONVERTLIBINTERFACES;
using System.Security.Cryptography;




namespace peredacha
{
    public struct Zveno
    {
        public double[] T;
        public double[] B3;
        public double[] H;
        public double[] D0;
        public double[] D1;
        public double[] S;

        public Zveno(out double[] t, out double[] b3, out double[] h, out double[] d0, out double[] d1, out double[] s)
        {
            t = new double[5] { 50, 60, 70, 80, 90 };
            b3 = new double[5] { 35, 45, 50, 60, 70 };
            h = new double[5] { 38, 45, 55, 60, 70 };
            d0 = new double[5] { 18, 23, 28, 32, 36 };
            d1 = new double[5] { 22, 26, 32, 36, 40 };
            s = new double[5] { 4.5, 6, 6, 6, 7 };
            T = t;
            B3 = b3;
            H = h;
            D0 = d0;
            D1 = d1;
            S = s;
        }     
    }

public partial class Form1 : Form
    {
        private Zveno zvenoInstance;
        private KompasObject kompas = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            zvenoInstance = new Zveno(out double[] t, out double[] b3, out double[] h, out double[] d0, out double[] d1, out double[] s);
            int number = Convert.ToInt32(textBox1.Text);
            for (int z = 0; z < 5; z++)
            {
                if (number == zvenoInstance.T[z])
                {       
                    if (kompas == null)
                    {
                        Type t1 = Type.GetTypeFromProgID("KOMPAS.Application.5");
                        kompas = (KompasObject)Activator.CreateInstance(t1);
                    }
                    if (kompas != null)
                    {
                        kompas.Visible = true;
                        kompas.ActivateControllerAPI();
                    }

                    if (kompas != null)
                    {
                        ksDocument3D iDocument3D = (ksDocument3D)kompas.Document3D();
                        iDocument3D.Create(false, true);
                        ksPart iPart = (ksPart)iDocument3D.GetPart((short)Part_Type.pTop_Part);
                        if (iPart != null)
                        {
                            //эскиз 1
                            ksEntity planeXOY = (ksEntity)iPart.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
                            ksEntity iSketch = (ksEntity)iPart.NewEntity((short)Obj3dType.o3d_sketch);

                            ksSketchDefinition iDefinitionSketch = (ksSketchDefinition)iSketch.GetDefinition();
                            iDefinitionSketch.SetPlane(planeXOY);
                            iSketch.Create();

                            ksDocument2D iDocument2D = (ksDocument2D)iDefinitionSketch.BeginEdit();

                            iDocument2D.ksArcByAngle(0, 0, (zvenoInstance.H[z])/2, 60, 300, 1, 1);
                            iDocument2D.ksArcByAngle(zvenoInstance.T[z], 0, (zvenoInstance.H[z]) / 2, 120, 240, -1, 1);
                            iDocument2D.ksArcByAngle(zvenoInstance.T[z]/2, -(((Math.Sqrt(3)) / 2) * zvenoInstance.T[z]), zvenoInstance.T[z] - (zvenoInstance.H[z] / 2), 60, 120, 1, 1);
                            iDocument2D.ksArcByAngle(zvenoInstance.T[z]/2, (((Math.Sqrt(3)) / 2) * zvenoInstance.T[z]), zvenoInstance.T[z] - (zvenoInstance.H[z] / 2), 240, 300, 1, 1);
                            //iDocument2D.ksArcByAngle(25, (((Math.Sqrt(3)) / 2) * 50), 50 - 19, 240, 300, 1, 1);
                            iDocument2D.ksCircle(0, 0, zvenoInstance.D0[z] / 2, 1);
                            iDocument2D.ksCircle(zvenoInstance.T[z], 0, zvenoInstance.D0[z] / 2, 1);    

                            iDefinitionSketch.EndEdit();

                            //выдавливание
                            ksEntity entityExtr = (ksEntity)iPart.NewEntity((short)Obj3dType.o3d_bossExtrusion);
                            ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition();

                            ksExtrusionParam extrProp = (ksExtrusionParam)extrusionDef.ExtrusionParam();
                            ksThinParam thinProp = (ksThinParam)extrusionDef.ThinParam();
                            if (extrProp != null && thinProp != null)
                            {
                                extrusionDef.SetSketch(iSketch); 

                                extrProp.direction = (short)Direction_Type.dtNormal;
                                extrProp.typeNormal = (short)End_Type.etBlind;
                                extrProp.depthNormal = zvenoInstance.S[z];
                                entityExtr.Create();
                            }

                            //эскиз 2
                            ksEntity entitySketch2 = (ksEntity)iPart.NewEntity((short)Obj3dType.o3d_sketch);
                            ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
                            sketchDef2.SetPlane(planeXOY);
                            entitySketch2.Create();

                            ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit();
                            sketchEdit2.ksCircle(0, 0, zvenoInstance.D0[z] / 2, 1);
                            sketchEdit2.ksCircle(zvenoInstance.T[z], 0, zvenoInstance.D0[z] / 2, 1);
                            sketchEdit2.ksCircle(0, 0, zvenoInstance.D1[z] / 2, 1);
                            sketchEdit2.ksCircle(zvenoInstance.T[z], 0, zvenoInstance.D1[z] / 2, 1);
                            sketchDef2.EndEdit();

                            // выдавливание
                            ksEntity entityBossExtr = (ksEntity)iPart.NewEntity((short)Obj3dType.o3d_bossExtrusion);

                            ksBossExtrusionDefinition bossExtrDef = (ksBossExtrusionDefinition)entityBossExtr.GetDefinition();
                            if (bossExtrDef != null)
                            {
                                ksExtrusionParam extrProp1 = (ksExtrusionParam)bossExtrDef.ExtrusionParam();
                                ksThinParam thinProp1 = (ksThinParam)bossExtrDef.ThinParam();
                                if (extrProp1 != null && thinProp1 != null)
                                {
                                    bossExtrDef.SetSketch(entitySketch2);

                                    extrProp1.direction = (short)Direction_Type.dtReverse;
                                    extrProp1.typeNormal = (short)End_Type.etBlind;
                                    extrProp1.depthReverse = zvenoInstance.B3[z] / 2;
                                    entityBossExtr.Create();
                                }
                            }
                            //смещенная плоскость
                            ksEntity EnPlaneOffset = iPart.NewEntity((short)Obj3dType.o3d_planeOffset);
                            ksPlaneOffsetDefinition PlaneOffsetDef = EnPlaneOffset.GetDefinition();
                            PlaneOffsetDef.direction = true;
                            PlaneOffsetDef.offset = -zvenoInstance.B3[z] / 2;
                            PlaneOffsetDef.SetPlane(iPart.GetDefaultEntity((short)Obj3dType.o3d_planeXOY));
                            EnPlaneOffset.Create();

                            //зеркальное отражение
                            ksEntity EnMirrorOp = iPart.NewEntity((short)Obj3dType.o3d_mirrorOperation);
                            ksMirrorCopyDefinition MirrorCopyDef = EnMirrorOp.GetDefinition();
                            MirrorCopyDef.SetPlane(EnPlaneOffset);
                            ksEntityCollection EnColl = MirrorCopyDef.GetOperationArray();
                            EnColl.Clear();
                            EnColl.Add(entityBossExtr);
                            EnColl.Add(entityExtr);
                            EnMirrorOp.Create();

                            ksEntityCollection ksEntityCollection = (ksEntityCollection)iPart.EntityCollection((short)Obj3dType.o3d_face);
                            Dictionary<string, int> nameCounter = new Dictionary<string, int>();

                            for (int i = 0; i < ksEntityCollection.GetCount(); i++)
                            {
                                ksEntity part1 = ksEntityCollection.GetByIndex(i);
                                ksFaceDefinition def = part1.GetDefinition();

                                if (def.GetOwnerEntity() == entityExtr)
                                {
                                    if (def.IsCylinder())
                                    {
                                        double h1, r;
                                        def.GetCylinderParam(out h1, out r);

                                        if (r == zvenoInstance.D0[z] / 2)
                                        {
                                            string baseName = "Cylinder";

                                            if (!nameCounter.ContainsKey(baseName))
                                            {
                                                nameCounter[baseName] = 1;
                                            }

                                            int counter = nameCounter[baseName];
                                            part1.name = $"{baseName}_{counter}";
                                            part1.Update();

                                            nameCounter[baseName]++;
                                        }
                                    }
                                }
                            }
                            ksEntityCollection = (ksEntityCollection)iPart.EntityCollection((short)Obj3dType.o3d_face);
                            for (int q = 0; q < ksEntityCollection.GetCount(); q++)
                            {
                                ksEntity part11 = ksEntityCollection.GetByIndex(q);
                                ksFaceDefinition def1 = part11.GetDefinition();
                                if (def1.GetOwnerEntity() == entityExtr)
                                {
                                    if (def1.IsPlanar())
                                    {
                                        ksEdgeCollection col = def1.EdgeCollection();
                                        for (int k = 0; k < col.GetCount(); k++)
                                        {
                                            ksEdgeDefinition d = col.GetByIndex(k);
                                            if (d.IsCircle())
                                            {
                                                ksVertexDefinition p1 = d.GetVertex(true);
                                                double x1, y1, z1;
                                                p1.GetPoint(out x1, out y1, out z1);
                                                if (x1 == zvenoInstance.D0[z] / 2 && y1 == 0)
                                                {
                                                    part11.name = ("Panel_1");
                                                    part11.Update();
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }

                            }
                        }

                        iDocument3D.SaveAs("C:\\с\\Part11.m3d");

                        //2 деталь
                        ksDocument3D iDocument3D1 = (ksDocument3D)kompas.Document3D();
                        iDocument3D1.Create(false, true);
                        ksPart iPart1 = (ksPart)iDocument3D1.GetPart((short)Part_Type.pTop_Part);
                        if (iPart1 != null)
                        {
                            // эскиз 3
                            ksEntity planeXOY = (ksEntity)iPart1.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
                            ksEntity iSketch = (ksEntity)iPart1.NewEntity((short)Obj3dType.o3d_sketch);

                            ksSketchDefinition iDefinitionSketch3 = (ksSketchDefinition)iSketch.GetDefinition();
                            iDefinitionSketch3.SetPlane(planeXOY);
                            iSketch.Create();
                            ksDocument2D iDocument2D1 = (ksDocument2D)iDefinitionSketch3.BeginEdit();

                            iDocument2D1.ksArcByAngle(0, 0, (zvenoInstance.H[z]) / 2, 60, 300, 1, 1);
                            iDocument2D1.ksArcByAngle(zvenoInstance.T[z], 0, (zvenoInstance.H[z]) / 2, 120, 240, -1, 1);
                            iDocument2D1.ksArcByAngle(zvenoInstance.T[z] / 2, -(((Math.Sqrt(3)) / 2) * zvenoInstance.T[z]), zvenoInstance.T[z] - (zvenoInstance.H[z] / 2), 60, 120, 1, 1);
                            iDocument2D1.ksArcByAngle(zvenoInstance.T[z] / 2, (((Math.Sqrt(3)) / 2) * zvenoInstance.T[z]), zvenoInstance.T[z] - (zvenoInstance.H[z] / 2), 240, 300, 1, 1);

                            iDefinitionSketch3.EndEdit();

                            ksEntity entityExtr3 = (ksEntity)iPart1.NewEntity((short)Obj3dType.o3d_bossExtrusion);

                            // выдавливание
                            ksBossExtrusionDefinition extrusionDef3 = (ksBossExtrusionDefinition)entityExtr3.GetDefinition();

                            ksExtrusionParam extrProp3 = (ksExtrusionParam)extrusionDef3.ExtrusionParam();
                            ksThinParam thinProp3 = (ksThinParam)extrusionDef3.ThinParam();
                            if (extrProp3 != null && thinProp3 != null)
                            {
                                extrusionDef3.SetSketch(iSketch); 

                                extrProp3.direction = (short)Direction_Type.dtNormal;
                                extrProp3.typeNormal = (short)End_Type.etBlind;
                                extrProp3.depthNormal = zvenoInstance.S[z];
                                entityExtr3.Create();
                            }

                            //смещенная плоскость
                            ksEntity EnPlaneOffset5 = iPart1.NewEntity((short)Obj3dType.o3d_planeOffset);
                            ksPlaneOffsetDefinition PlaneOffsetDef5 = EnPlaneOffset5.GetDefinition();
                            PlaneOffsetDef5.direction = true;
                            PlaneOffsetDef5.offset = zvenoInstance.S[z];
                            PlaneOffsetDef5.SetPlane(iPart.GetDefaultEntity((short)Obj3dType.o3d_planeXOY));
                            EnPlaneOffset5.Create();

                            ////эскиз 4
                            ksEntity entitySketch4 = (ksEntity)iPart1.NewEntity((short)Obj3dType.o3d_sketch);
                            ksSketchDefinition sketchDef4 = (ksSketchDefinition)entitySketch4.GetDefinition();
                            sketchDef4.SetPlane(planeXOY);
                            entitySketch4.Create();

                            ksDocument2D sketchEdit4 = (ksDocument2D)sketchDef4.BeginEdit();
                            sketchEdit4.ksCircle(0, 0, zvenoInstance.D0[z]/2, 1);
                            sketchEdit4.ksCircle(zvenoInstance.T[z], 0, zvenoInstance.D0[z]/2, 1);
                            sketchDef4.EndEdit();

                            //// выдавливание
                            ksEntity entityBossExtr4 = (ksEntity)iPart1.NewEntity((short)Obj3dType.o3d_bossExtrusion);

                            ksBossExtrusionDefinition bossExtrDef4 = (ksBossExtrusionDefinition)entityBossExtr4.GetDefinition();
                            if (bossExtrDef4 != null)
                            {
                                ksExtrusionParam extrProp4 = (ksExtrusionParam)bossExtrDef4.ExtrusionParam();
                                ksThinParam thinProp4 = (ksThinParam)bossExtrDef4.ThinParam();
                                if (extrProp4 != null && thinProp4 != null)
                                {
                                    bossExtrDef4.SetSketch(entitySketch4);

                                    extrProp4.direction = (short)Direction_Type.dtReverse;
                                    extrProp4.typeNormal = (short)End_Type.etBlind;
                                    extrProp4.depthReverse = (zvenoInstance.B3[z] / 2) + zvenoInstance.S[z];
                                    entityBossExtr4.Create();
                                }
                            }

                            //эскиз 5
                            ksEntity entitySketch6 = (ksEntity)iPart1.NewEntity((short)Obj3dType.o3d_sketch);
                            ksSketchDefinition sketchDef6 = (ksSketchDefinition)entitySketch6.GetDefinition();
                            sketchDef6.SetPlane(EnPlaneOffset5);
                            entitySketch6.Create();

                            ksDocument2D sketchEdit6 = (ksDocument2D)sketchDef6.BeginEdit();
                            sketchEdit6.ksCircle(0, 0, zvenoInstance.D1[z]/2, 1);
                            sketchEdit6.ksCircle(zvenoInstance.T[z], 0, zvenoInstance.D1[z]/2, 1);
                            sketchDef6.EndEdit();

                            //выдавливание
                            ksEntity entityBossExtr6 = (ksEntity)iPart1.NewEntity((short)Obj3dType.o3d_bossExtrusion);

                            ksBossExtrusionDefinition bossExtrDef6 = (ksBossExtrusionDefinition)entityBossExtr6.GetDefinition();
                            if (bossExtrDef6 != null)
                            {
                                ksExtrusionParam extrProp6 = (ksExtrusionParam)bossExtrDef6.ExtrusionParam();
                                ksThinParam thinProp6 = (ksThinParam)bossExtrDef6.ThinParam();
                                if (extrProp6 != null && thinProp6 != null)
                                {
                                    bossExtrDef6.SetSketch(entitySketch6);

                                    extrProp6.direction = (short)Direction_Type.dtReverse;
                                    extrProp6.typeNormal = (short)End_Type.etBlind;
                                    extrProp6.depthReverse = -2;
                                    entityBossExtr6.Create();
                                }
                            }

                            ksEntityCollection ksEntityCollection = (ksEntityCollection)iPart1.EntityCollection((short)Obj3dType.o3d_face);
                            Dictionary<string, int> nameCounter = new Dictionary<string, int>();

                            for (int i = 0; i < ksEntityCollection.GetCount(); i++)
                            {
                                ksEntity part2 = ksEntityCollection.GetByIndex(i);
                                ksFaceDefinition def = part2.GetDefinition();

                                if (def.GetOwnerEntity() == entityBossExtr4)
                                {
                                    if (def.IsCylinder())
                                    {
                                        double h2, r;
                                        def.GetCylinderParam(out h2, out r);

                                        if (r == (zvenoInstance.D0[z] / 2))
                                        {
                                            string baseName = "Cylinder_2";

                                            if (!nameCounter.ContainsKey(baseName))
                                            {
                                                nameCounter[baseName] = 1;
                                            }

                                            int counter = nameCounter[baseName];
                                            part2.name = $"{baseName}_{counter}";
                                            part2.Update();

                                            nameCounter[baseName]++;
                                        }
                                    }
                                }
                            }

                            ksEntityCollection = (ksEntityCollection)iPart1.EntityCollection((short)Obj3dType.o3d_face);
                            for (int q = 0; q < ksEntityCollection.GetCount(); q++)
                            {
                                ksEntity part12 = ksEntityCollection.GetByIndex(q);
                                ksFaceDefinition def1 = part12.GetDefinition();
                                if (def1.GetOwnerEntity() == entityBossExtr6)
                                {
                                    if (def1.IsPlanar())
                                    {
                                        ksEdgeCollection col = def1.EdgeCollection();
                                        for (int k = 0; k < col.GetCount(); k++)
                                        {
                                            ksEdgeDefinition d = col.GetByIndex(k);
                                            if (d.IsCircle())
                                            {
                                                ksVertexDefinition p1 = d.GetVertex(true);
                                                double x1, y1, z1;
                                                p1.GetPoint(out x1, out y1, out z1);
                                                if (x1 == zvenoInstance.D1[z] / 2 && z1 == 2 + zvenoInstance.S[z])
                                                {
                                                    part12.name = ("Panel_2");
                                                    part12.Update();
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            //скругление 1
                            ksEntity Fillet = iPart1.NewEntity((short)Obj3dType.o3d_fillet);
                            ksFilletDefinition FilletDef = Fillet.GetDefinition();
                            FilletDef.radius = 2;
                            FilletDef.tangent = false;

                            ksEntityCollection EnColPart = iPart1.EntityCollection((short)Obj3dType.o3d_edge);
                            ksEntityCollection EnColFillet = FilletDef.array();
                            EnColFillet.Clear();

                            EnColPart.SelectByPoint(zvenoInstance.D1[z] / 2, 0, 2 + zvenoInstance.S[z]);
                            EnColFillet.Add(EnColPart.GetByIndex(0));
                            Fillet.Create();

                            //скругление 2
                            ksEntity Fillet2 = iPart1.NewEntity((short)Obj3dType.o3d_fillet);
                            ksFilletDefinition FilletDef2 = Fillet2.GetDefinition();
                            FilletDef2.radius = 2;
                            FilletDef2.tangent = false;

                            ksEntityCollection EnColPart2 = iPart1.EntityCollection((short)Obj3dType.o3d_edge);
                            ksEntityCollection EnColFillet2 = FilletDef2.array();
                            EnColFillet2.Clear();

                            EnColPart2.SelectByPoint(zvenoInstance.T[z] + zvenoInstance.D1[z] / 2, 0, 2 + zvenoInstance.S[z]);
                            EnColFillet2.Add(EnColPart2.GetByIndex(0));
                            Fillet2.Create();

                            //смещенная плоскость
                            ksEntity EnPlaneOffset1 = iPart1.NewEntity((short)Obj3dType.o3d_planeOffset);
                            ksPlaneOffsetDefinition PlaneOffsetDef1 = EnPlaneOffset1.GetDefinition();
                            PlaneOffsetDef1.direction = true;
                            PlaneOffsetDef1.offset = -((zvenoInstance.B3[z] / 2) + zvenoInstance.S[z]);
                            PlaneOffsetDef1.SetPlane(iPart.GetDefaultEntity((short)Obj3dType.o3d_planeXOY));
                            EnPlaneOffset1.Create();

                            //зеркальное отражение
                            ksEntity EnMirrorOp1 = iPart1.NewEntity((short)Obj3dType.o3d_mirrorOperation);
                            ksMirrorCopyDefinition MirrorCopyDef1 = EnMirrorOp1.GetDefinition();
                            MirrorCopyDef1.SetPlane(EnPlaneOffset1);
                            ksEntityCollection EnColl1 = MirrorCopyDef1.GetOperationArray();
                            EnColl1.Clear();
                            EnColl1.Add(entityBossExtr4);
                            EnColl1.Add(entityExtr3);
                            EnColl1.Add(entityBossExtr6);
                            EnColl1.Add(Fillet);
                            EnColl1.Add(Fillet2);
                            EnMirrorOp1.Create();
                                                       
                        }
                        iDocument3D1.SaveAs("C:\\с\\Part12.m3d");
                                                
                    }
                    ksDocument3D iDocument3Dsb = (ksDocument3D)kompas.Document3D();
                    iDocument3Dsb.Create(false, false);
                    ksPart Partsb = (ksPart)iDocument3Dsb.GetPart((short)Part_Type.pTop_Part);
                    iDocument3Dsb.SetPartFromFile("C:\\с\\Part11.m3d", Partsb, false);
                    iDocument3Dsb.SetPartFromFile("C:\\с\\Part12.m3d", Partsb, false);

                    ksPartCollection partColl = iDocument3Dsb.PartCollection(true);

                    Partsb = partColl.GetByIndex(0);
                    ksEntityCollection PartCol = Partsb.EntityCollection((short)Obj3dType.o3d_face);
                    ksEntity cylinder1 = PartCol.GetByName("Cylinder_1", true, true);
                    ksEntity cylinder2 = PartCol.GetByName("Cylinder_2", true, true);
                    ksEntity panel_1 = PartCol.GetByName("Panel_1", true, true);

                    Partsb = partColl.GetByIndex(1);
                    PartCol = Partsb.EntityCollection((short)Obj3dType.o3d_face);
                    ksEntity cylinder12 = PartCol.GetByName("Cylinder_2_1", true, true);
                    ksEntity cylinder22 = PartCol.GetByName("Cylinder_2_2", true, true);
                    ksEntity panel_2 = PartCol.GetByName("Panel_2", true, true);

                    
                    iDocument3Dsb.AddMateConstraint(4, cylinder1, cylinder12, 1, 1, 0);
                    iDocument3Dsb.AddMateConstraint(5, panel_1, panel_2, 1, 0, 2 + zvenoInstance.S[z]);


                }
            }

        }
    }
}
