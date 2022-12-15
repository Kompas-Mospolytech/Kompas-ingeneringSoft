using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Ks3DModels
{
    public partial class Form1 : Form
    {
		private KompasObject kompas;
        private ksDocument3D doc;
        private ksDocument3D ksDoc3d;
        private ksDocument3D ksShaftDoc3d;
        private Dictionary<string, ksDocument3D> Doc3d = new Dictionary<string, ksDocument3D>();
        private ksDocument3D modelOnDelete;

		public double L1 = 35, L2 = 30, L3 = 9, R1 = 8, R2 = 22.5;
		public Form1()
		{
			InitializeComponent();

            this.comboBox_L1.Text = System.Convert.ToString(L1);
            this.comboBox_L1.Items.AddRange(new object[] { "45", "55", "65", "75", "85", "95", "105" });
            this.comboBox_L2.Text = System.Convert.ToString(L2);
            this.comboBox_L2.Items.AddRange(new object[] { "40", "50", "60", "70", "80", "90", "100" });
            this.comboBox_L3.Text = System.Convert.ToString(L3);
            this.comboBox_L3.Items.AddRange(new object[] { "12", "14", "16", "18", "20"});
            this.comboBox_R1.Text = System.Convert.ToString(R1);
            this.comboBox_R1.Items.AddRange(new object[] { "9", "10", "11", "12"});
            this.comboBox_R2.Text = System.Convert.ToString(R2);
            this.comboBox_R2.Items.AddRange(new object[] { "25", "30", "35", "40", "45", "50" });

        }
		// Удалить все объекты из текущего эскиза
		void ClearCurrentSketch(ksDocument2D sketchEdit)
		{
			// создаим итератор и удалим все существующие объекты в эскизе          
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter != null)
			{
				if (iter.ksCreateIterator(ldefin2d.ALL_OBJ, 0))
				{
					reference rf;
					if ((rf = iter.ksMoveIterator("F")) != 0)
					{
						// сместить указатель на первый элемент в списке
						// в цикле сместить указатель на следующий элемент в списке пока не дойдем до последнего
						do
						{
							if (sketchEdit.ksExistObj(rf) == 1)
								sketchEdit.ksDeleteObj(rf); // если объект существует удалить его 
						}
						while ((rf = iter.ksMoveIterator("N")) != 0);
					}
					iter.ksDeleteIterator();    // удалим итератор
				}
			}
		}

        private void comboBox_L1_TextChanged(object sender, EventArgs e)
        {
           L1 = System.Convert.ToDouble(comboBox_L1.Text);
        }

        private void comboBox_L2_TextChanged(object sender, EventArgs e)
        {
            L2 = System.Convert.ToDouble(comboBox_L2.Text);
        }

        private void comboBox_L3_TextChanged(object sender, EventArgs e)
        {
            L3 = System.Convert.ToDouble(comboBox_L3.Text);
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox_R1_TextChanged(object sender, EventArgs e)
        {
            R1 = System.Convert.ToDouble(comboBox_R1.Text);
        }

        private void comboBox_L3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox_R2_TextChanged(object sender, EventArgs e)
        {
            R2 = System.Convert.ToDouble(comboBox_R2.Text);
        }

        void ClearCurrentModel(ksDocument3D model, ksPart part)
		{
			IntPtr rf;
			rf = part.GetFeature();
			model.ksDeleteObj((int)rf); 

		}


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            }
            catch
            {
                kompas = (KompasObject)Activator.CreateInstance(Type.GetTypeFromProgID("KOMPAS.Application.5"));
            }
            if (kompas == null)
                return;
            kompas.Visible = true;
			ksDocument3D idoc3d = kompas.Document3D();

			doc = (ksDocument3D)kompas.ActiveDocument3D();
			if (doc == null || doc.reference == 0)
			{
				doc = (ksDocument3D)kompas.Document3D();
				doc.Create(true, true);

				doc.author = "Artem";
				doc.comment = "3DModel";
				doc.UpdateDocumentParam();
			}

			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part); // новый компонент
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// интерфейс свойств эскиза
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// получим интерфейс базовой плоскости XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);  // установим плоскость XOY базовой для эскиза
						sketchDef.angle = 45;           // угол поворота эскиза
						entitySketch.Create();          // создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						// введем новый эскиз - квадрат
						sketchEdit.ksLineSeg(50, 50, -50, 50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1);

						sketchEdit.ksLineSeg(50, -50, 50, 50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50, 50, 1);
						sketchDef.EndEdit();    // завершение редактирования эскиза

						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// интерфейс свойств базовой операции выдавливания
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition(); // интерфейс базовой операции выдавливания
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;         // направление выдавливания
								extrusionDef.SetSideParam(true, // прямое направление
									(short)End_Type.etBlind,    // строго на глубину
									200, 0, false);
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10); // тонкая стенка в два направления
								extrusionDef.SetSketch(entitySketch);   // эскиз операции выдавливания
								entityExtr.Create();                    // создать операцию

								kompas.ksMessage("Изменим параметры операции выдавливания");

								extrusionDef.SetSideParam(false,    // обратное направление
									(short)End_Type.etBlind,        // строго на глубину
									150, 0, false);
								extrusionDef.directionType = (short)Direction_Type.dtBoth; // направление выдавливания dtBoth - в оба направления
								entityExtr.Update();    // обновить параметры

								kompas.ksMessage("Поменяем эскиз");

								// режим редактирования эскиза
								sketchEdit = (ksDocument2D)sketchDef.BeginEdit();

								// создаим итератор и удалим все существующие объекты в эскизе
								ClearCurrentSketch(sketchEdit);
								// введем в эскиз окружность
								sketchEdit.ksCircle(0, 0, 100, 1);

								sketchDef.EndEdit();    // завершение редактирования эскиза
								entitySketch.Update();  // обновить параметры эскиза
								entityExtr.Update();    // обновить параметры операции выдавливания

								kompas.ksMessage("Приклеем выдавливанием");

								// создадим новый эскиз
								ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
								if (entitySketch2 != null)
								{
									// интерфейс свойств эскиза
									ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
									if (sketchDef2 != null)
									{
										sketchDef2.SetPlane(basePlane); // установим плоскость
										sketchDef2.angle = 45;          // повернем эскиз на 45 град.
										entitySketch2.Create();         // создадим эскиз

										// интерфейс редактора эскиза
										ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit();
										sketchEdit2.ksCircle(0, 0, 150, 1);
										sketchDef2.EndEdit();   // завершение редактирования эскиза

										// приклеим выдавливанием
										ksEntity entityBossExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
										if (entityBossExtr != null)
										{
											ksBossExtrusionDefinition bossExtrDef = (ksBossExtrusionDefinition)entityBossExtr.GetDefinition();
											if (bossExtrDef != null)
											{
												ksExtrusionParam extrProp = (ksExtrusionParam)bossExtrDef.ExtrusionParam(); // интерфейс структуры параметров выдавливания
												ksThinParam thinProp = (ksThinParam)bossExtrDef.ThinParam();      // интерфейс структуры параметров тонкой стенки
												if (extrProp != null && thinProp != null)
												{
													bossExtrDef.SetSketch(entitySketch2); // эскиз операции выдавливания

													extrProp.direction = (short)Direction_Type.dtNormal;      // направление выдавливания (прямое)
													extrProp.typeNormal = (short)End_Type.etBlind;      // тип выдавливания (строго на глубину)
													extrProp.depthNormal = 100;         // глубина выдавливания

													thinProp.thin = false;              // без тонкой стенки

													entityBossExtr.Create();                // создадим операцию
												}
											}
										}

										// создадим новый эскиз
										ksEntity entitySketch3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
										if (entitySketch3 != null)
										{
											// интерфейс свойств эскиза
											ksSketchDefinition sketchDef3 = (ksSketchDefinition)entitySketch3.GetDefinition();
											if (sketchDef3 != null)
											{
												sketchDef3.SetPlane(basePlane); // установим плоскость
												sketchDef3.angle = 45;          // повернем эскиз на 45 град.
												entitySketch3.Create();         // создадим эскиз

												// интерфейс редактора эскиза
												ksDocument2D sketchEdit3 = (ksDocument2D)sketchDef3.BeginEdit();
												// введем новый эскиз - квадрат
												sketchEdit3.ksLineSeg(50, 50, -50, 50, 1);
												sketchEdit3.ksLineSeg(50, -50, -50, -50, 1);
												sketchEdit3.ksLineSeg(50, -50, 50, 50, 1);
												sketchEdit3.ksLineSeg(-50, -50, -50, 50, 1);
												sketchDef3.EndEdit();   // завершение редактирования эскиза

												// вырежим выдавливанием
												ksEntity entityCutExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
												if (entityCutExtr != null)
												{
													ksCutExtrusionDefinition cutExtrDef = (ksCutExtrusionDefinition)entityCutExtr.GetDefinition();
													if (cutExtrDef != null)
													{
														cutExtrDef.SetSketch(entitySketch3);    // установим эскиз операции
														cutExtrDef.directionType = (short)Direction_Type.dtNormal; //прямое направление
														cutExtrDef.SetSideParam(true, (short)End_Type.etBlind, 50, 0, false);
														cutExtrDef.SetThinParam(false, 0, 0, 0);
													}

													entityCutExtr.Create(); // создадим операцию вырезание выдавливанием
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}

        private void button2_Click(object sender, EventArgs e)
        {
            try // проверяется, открыт ли компас
            {
				kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            }
            catch // если не открыт то создаётся инстанс компаса
            {
				kompas = (KompasObject)Activator.CreateInstance(Type.GetTypeFromProgID("KOMPAS.Application.5"));
            }
			if (kompas == null)
				return;
			kompas.Visible = true; // видимость компаса
			ksDoc3d = kompas.Document3D(); // создание 3д документа

			ksDoc3d.Create(false, true); // false - видимый режим, true - деталь
			ksDoc3d = kompas.ActiveDocument3D(); // указатель на интерфейс 3д модели 
			ksDoc3d.author = "Artem";

			ksPart part = ksDoc3d.GetPart((int)Part_Type.pTop_Part); // новый компонент
			ksEntity ksScetchE = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча



			SketchDefinition ksScetchDef = ksScetchE.GetDefinition(); // интерфейс свойств эскиза
			// получим интерфейс базовой плоскости XOZ
			ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
			ksScetchDef.SetPlane(basePlane);  // установим плоскость XOY базовой для эскиза
			ksScetchE.Create();          // создадим эскиз
			ksDocument2D Scetch2D = (ksDocument2D)ksScetchDef.BeginEdit();
			Scetch2D.ksLineSeg(-L1, R2, L1, R2, 1); // создание отрезка xc, yc - координаты начала и конца отрезка, style - стиль линии.
			Scetch2D.ksLineSeg(L1, R2, L1, R2+12.5, 1);
			Scetch2D.ksLineSeg(L3, R2 + 12.5, L1, R2 + 12.5, 1);
			Scetch2D.ksLineSeg(L3, R2 + 12.5, L3, R2 + 65.5, 1);
			Scetch2D.ksLineSeg(L3, R2+65.5, L2, R2 + 65.5, 1);
			Scetch2D.ksLineSeg(L2, R2 + 65.5, L2, R2 + 105.5, 1);

			Scetch2D.ksLineSeg(-L2, R2 + 105.5, L2, R2 + 105.5, 1);

			Scetch2D.ksLineSeg(-L1, R2, -L1, R2 + 12.5, 1);
			Scetch2D.ksLineSeg(-L3, R2 + 12.5, -L1, R2 + 12.5, 1);
			Scetch2D.ksLineSeg(-L3, R2 + 12.5, -L3, R2 + 65.5, 1);
			Scetch2D.ksLineSeg(-L3, R2 + 65.5, -L2, R2 + 65.5, 1);
			Scetch2D.ksLineSeg(-L2, R2 + 65.5, -L2, R2 + 105.5, 1);

			Scetch2D.ksLineSeg(-30, 0, 30, 0, 3);


			ksScetchDef.EndEdit(); // заканчивает редактирование эскиза

			ksEntity bossRot = part.NewEntity((short)Obj3dType.o3d_bossRotated); // сущность для вращения
			BossRotatedDefinition RotDef = bossRot.GetDefinition(); // интерфейс настроек вращения
			RotDef.SetSideParam(true,360); // ture - прямое выдавливание 360 градусов
			RotDef.directionType = (short)ksDirectionTypeEnum.dtNormal;         // направление выдавливания
			RotDef.SetSketch(ksScetchE);   // эскиз операции выдавливания
			bossRot.Create();  // создать операцию



			ksEntity ksScetchE2 = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

			SketchDefinition ksScetchDef2 = ksScetchE2.GetDefinition(); // интерфейс свойств эскиза
																	  // получим интерфейс базовой плоскости XOZ
			ksEntity basePlane2 = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);


			ksEntity basePlaneOffset = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);

			PlaneOffsetDefinition offsetPlaneDef = basePlaneOffset.GetDefinition();
			offsetPlaneDef.direction = true;
			offsetPlaneDef.offset = 70;
			offsetPlaneDef.SetPlane(basePlane2);
			basePlaneOffset.Create();


			ksScetchDef2.SetPlane(basePlaneOffset);  // установим плоскость XOY базовой для эскиза
			ksScetchE2.Create();          // создадим эскиз
			ksDocument2D Scetch2D_2 = (ksDocument2D)ksScetchDef2.BeginEdit();
			Scetch2D_2.ksLineSeg(-3, R2+94.5, 3, R2 + 94.5, 1); // создание отрезка xc, yc - координаты начала и конца отрезка, style - стиль линии.
			Scetch2D_2.ksLineSeg(-8.54, R2 + 113.5, 8.54, R2 + 113.5, 1);
			Scetch2D_2.ksLineSeg(-3, R2 + 94.5, -8.54, R2 + 113.5, 1);
			Scetch2D_2.ksLineSeg(3, R2 + 94.5, 8.54, R2 + 113.5, 1);

			ksScetchDef2.EndEdit(); // заканчивает редактирование эскиза

			ksEntity cutExtr = part.NewEntity((short)Obj3dType.o3d_cutExtrusion); 
			CutExtrusionDefinition cutRef = cutExtr.GetDefinition();
			cutRef.SetSideParam(true, (short)End_Type.etThroughAll);
			cutRef.directionType = (short)ksDirectionTypeEnum.dtNormal;         
			cutRef.SetSketch(ksScetchE2);  
			cutExtr.Create();  

			ksEntity circCopy = part.NewEntity((short)Obj3dType.o3d_circularCopy);
			CircularCopyDefinition circCopyDef = circCopy.GetDefinition();
			ksEntity baseAxis = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_axisOZ);
			circCopyDef.SetAxis(baseAxis);
			circCopyDef.SetCopyParamAlongDir(31,12,false,false);
			ksEntityCollection EntityCollection = circCopyDef.GetOperationArray();
			EntityCollection.Clear();
			EntityCollection.Add(cutExtr);
			circCopy.Create();
            EntityCollection.Clear();



            ksEntity ksScetchE3 = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

            SketchDefinition ksScetchDef3 = ksScetchE3.GetDefinition(); // интерфейс свойств эскиза
            ksScetchDef3.SetPlane(basePlaneOffset);  // установим плоскость XOY базовой для эскиза
            ksScetchE3.Create();          // создадим эскиз
            ksDocument2D Scetch2D_3 = (ksDocument2D)ksScetchDef3.BeginEdit();
            Scetch2D_3.ksCircle(0,R2+33.5,R1, 1);

            ksScetchDef3.EndEdit(); // заканчивает редактирование эскиза

            ksEntity cutExtr2 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            CutExtrusionDefinition cutRef2 = cutExtr2.GetDefinition();
            cutRef2.SetSideParam(true, (short)End_Type.etThroughAll);
            cutRef2.directionType = (short)ksDirectionTypeEnum.dtNormal;
            cutRef2.SetSketch(ksScetchE3);
            cutExtr2.Create();

			ksEntity circCopy2 = part.NewEntity((short)Obj3dType.o3d_circularCopy);
			CircularCopyDefinition circCopyDef2 = circCopy2.GetDefinition();
			circCopyDef2.SetAxis(baseAxis);
			circCopyDef2.SetCopyParamAlongDir(4, 90, false, false);
			ksEntityCollection EntityCollection2 = circCopyDef2.GetOperationArray();
			EntityCollection2.Clear();
			EntityCollection2.Add(cutExtr2);
			circCopy2.Create();


            ksEntity ksScetchE4 = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

            SketchDefinition ksScetchDef4 = ksScetchE4.GetDefinition(); // интерфейс свойств эскиза
            ksScetchDef4.SetPlane(basePlaneOffset);  // установим плоскость XOY базовой для эскиза
            ksScetchE4.Create();          // создадим эскиз
            ksDocument2D Scetch2D_4 = (ksDocument2D)ksScetchDef4.BeginEdit();
			ksRectangleParam ksRecParam = kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
			ksRecParam.height = R2 + 1.65;
			ksRecParam.width = 14;
			ksRecParam.style = 1;
			ksRecParam.ang = 0;
            ksRecParam.x = -7;
			ksRecParam.y = 0;
			Scetch2D_4.ksRectangle(ksRecParam);


            ksScetchDef4.EndEdit(); // заканчивает редактирование эскиза

            ksEntity cutExtr3 = part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
            CutExtrusionDefinition cutRef3 = cutExtr3.GetDefinition();
            cutRef3.SetSideParam(true, (short)End_Type.etThroughAll);
            cutRef3.directionType = (short)ksDirectionTypeEnum.dtNormal;
            cutRef3.SetSketch(ksScetchE4);
            cutExtr3.Create();



            ksDoc3d.hideAllPlanes = true;
			

            /*            kompas.ksMessage("Шестерёнка построена");*/

            /*			ClearCurrentModel(ksDoc3d, part);*/

            /*			ksEntity ksScetchE2 = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

						SketchDefinition ksScetchDef2 = ksScetchE2.GetDefinition(); // интерфейс свойств эскиза

						ksScetchDef2.SetPlane(basePlane);  // установим плоскость XOY базовой для эскиза
						ksScetchE2.Create();          // создадим эскиз
						ksDocument2D Scetch2D_2 = (ksDocument2D)ksScetchDef2.BeginEdit();
						Scetch2D_2.ksCircle(0, 0, 70, 1);
						Scetch2D_2.ksCircle(0, 0, 45, 1);

						ksScetchDef2.EndEdit(); // заканчивает редактирование эскиза

						ksEntity bossExtr2 = part.NewEntity((short)Obj3dType.o3d_bossExtrusion); // сущность для выдавливания
						ksBossExtrusionDefinition ExtrDef2 = bossExtr2.GetDefinition(); // интерфейс настроек выдавливания


						ExtrDef2.SetSideParam(true, // прямое направление
							(short)End_Type.etBlind,    // строго на глубину
							70, 0, false);
						ExtrDef2.directionType = (short)Direction_Type.dtBoth;         // направление выдавливания
						ExtrDef2.SetSketch(ksScetchE2);   // эскиз операции выдавливания
						bossExtr2.Create();  // создать операцию



						ksEntity ksScetchE3 = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча

						SketchDefinition ksScetchDef3 = ksScetchE3.GetDefinition(); // интерфейс свойств эскиза

						ksScetchDef3.SetPlane(basePlane);  // установим плоскость XOY базовой для эскиза
						ksScetchE3.Create();          // создадим эскиз
						ksDocument2D Scetch2D_3 = (ksDocument2D)ksScetchDef3.BeginEdit();
						Scetch2D_3.ksCircle(0, 0, 70, 1);
						Scetch2D_3.ksCircle(0, 0, 176, 1);

						ksScetchDef3.EndEdit(); // заканчивает редактирование эскиза

						ksEntity bossExtr3 = part.NewEntity((short)Obj3dType.o3d_bossExtrusion); // сущность для выдавливания
						ksBossExtrusionDefinition ExtrDef3 = bossExtr3.GetDefinition(); // интерфейс настроек выдавливания

						ExtrDef3.directionType = (short)Direction_Type.dtBoth;         // направление выдавливания
						ExtrDef3.SetSideParam(true, // прямое направлениебьа
							(short)End_Type.etBlind,    // строго на глубину
							18, 0, false);
						ExtrDef3.SetSketch(ksScetchE3);   // эскиз операции выдавливания
						bossExtr3.Update();
						bossExtr3.Create();  // создать операцию*/




            /*			ksEntity CutExtr = part.NewEntity((short)Obj3dType.o3d_cutExtrusion); 
						CutExtrusionDefinition CutExtrDef = CutExtr.GetDefinition(); */

        }

        private void createShaft_Click(object sender, EventArgs e)
        {
            try // проверяется, открыт ли компас
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            }
            catch // если не открыт то создаётся инстанс компаса
            {
                kompas = (KompasObject)Activator.CreateInstance(Type.GetTypeFromProgID("KOMPAS.Application.5"));
            }
            if (kompas == null)
                return;
            kompas.Visible = true; // видимость компаса
            ksShaftDoc3d = kompas.Document3D(); // создание 3д документа

            ksShaftDoc3d.Create(false, true); // false - видимый режим, true - деталь
            ksShaftDoc3d = kompas.ActiveDocument3D(); // указатель на интерфейс 3д модели 
            ksShaftDoc3d.author = "Artem";
            ksPart part = ksShaftDoc3d.GetPart((int)Part_Type.pTop_Part); // новый компонент



            ksEntity ksScetchE = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча
            SketchDefinition ksScetchDef = ksScetchE.GetDefinition(); // интерфейс свойств эскиза
                                                                      // получим интерфейс базовой плоскости XOZ
            ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
            ksScetchDef.SetPlane(basePlane);  // установим плоскость XOY базовой для эскиза
            ksScetchE.Create();          // создадим эскиз
            ksDocument2D Scetch2D = (ksDocument2D)ksScetchDef.BeginEdit();
            Scetch2D.ksCircle(0, 0, 22.5, 1);


            ksScetchDef.EndEdit(); // заканчивает редактирование эскиза

            ksEntity bossExtr = part.NewEntity((short)Obj3dType.o3d_baseExtrusion); // сущность для выдавливания
            ksBaseExtrusionDefinition ExtrDef = bossExtr.GetDefinition(); // интерфейс настроек выдавливания
            ksExtrusionParam extrProp = (ksExtrusionParam)ExtrDef.ExtrusionParam();

            if (extrProp != null)
            {
                ExtrDef.SetSketch(ksScetchE); // эскиз операции выдавливания

                extrProp.direction = (short)Direction_Type.dtBoth;      // направление выдавливания (в обе стороны)
                extrProp.typeNormal = (short)End_Type.etBlind;      // тип выдавливания (строго на глубину)
                extrProp.depthNormal = 40;         // глубина выдавливания
                extrProp.depthReverse = 40;     // глубина выдавливания в обратном направлении
                bossExtr.Create();                // создадим операцию
            }


            ksEntity ksScetchE2 = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча
            SketchDefinition ksScetchDef2 = ksScetchE2.GetDefinition(); // интерфейс свойств эскиза
                                                                      // получим интерфейс базовой плоскости XOZ
            ksEntity basePlane2 = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
            ksScetchDef2.SetPlane(basePlane2);  // установим плоскость XOY базовой для эскиза
            ksScetchE2.Create();          // создадим эскиз
            ksDocument2D Scetch2D2 = (ksDocument2D)ksScetchDef2.BeginEdit();
            Scetch2D2.ksCircle(0, 0, 18, 1);


            ksScetchDef2.EndEdit(); // заканчивает редактирование эскиза

            ksEntity bossExtr2 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion); // сущность для выдавливания
            ksBaseExtrusionDefinition ExtrDef2 = bossExtr2.GetDefinition(); // интерфейс настроек выдавливания
            ksExtrusionParam extrProp2 = (ksExtrusionParam)ExtrDef2.ExtrusionParam();

            if (extrProp2 != null)
            {
                ExtrDef2.SetSketch(ksScetchE2); // эскиз операции выдавливания

                extrProp2.direction = (short)Direction_Type.dtBoth;      // направление выдавливания (в обе стороны)
                extrProp2.typeNormal = (short)End_Type.etBlind;      // тип выдавливания (строго на глубину)
                extrProp2.depthNormal = 0;         // глубина выдавливания
                extrProp2.depthReverse = 90;     // глубина выдавливания в обратном направлении
                bossExtr2.Create();                // создадим операцию
            }

            ksEntity ksScetchE3 = part.NewEntity((int)Obj3dType.o3d_sketch); // создание нового скетча
            SketchDefinition ksScetchDef3 = ksScetchE3.GetDefinition(); // интерфейс свойств эскиза
                                                                        // получим интерфейс базовой плоскости XOZ
            ksEntity basePlane3 = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
            ksScetchDef3.SetPlane(basePlane2);  // установим плоскость XOY базовой для эскиза
            ksScetchE3.Create();          // создадим эскиз
            ksDocument2D Scetch2D3 = (ksDocument2D)ksScetchDef3.BeginEdit();
            Scetch2D3.ksCircle(0, 0, 16, 1);


            ksScetchDef3.EndEdit(); // заканчивает редактирование эскиза

            ksEntity bossExtr3 = part.NewEntity((short)Obj3dType.o3d_baseExtrusion); // сущность для выдавливания
            ksBaseExtrusionDefinition ExtrDef3 = bossExtr3.GetDefinition(); // интерфейс настроек выдавливания
            ksExtrusionParam extrProp3 = (ksExtrusionParam)ExtrDef3.ExtrusionParam();

            if (extrProp3 != null)
            {
                ExtrDef3.SetSketch(ksScetchE3); // эскиз операции выдавливания

                extrProp3.direction = (short)Direction_Type.dtBoth;      // направление выдавливания (в обе стороны)
                extrProp3.typeNormal = (short)End_Type.etBlind;      // тип выдавливания (строго на глубину)
                extrProp3.depthNormal = 120;         // глубина выдавливания
                extrProp3.depthReverse = 0;     // глубина выдавливания в обратном направлении
                bossExtr3.Create();                // создадим операцию
            }

			ksShaftDoc3d.hideAllPlanes = true;

        }

        private void closeWheel_Click(object sender, EventArgs e)
        {
			if(ksDoc3d != null)
			{
				ksDoc3d.close();
            }

		}
        private void closeShaft_Click(object sender, EventArgs e)
        {
            if (ksShaftDoc3d != null)
            {
                ksShaftDoc3d.close();
            }
        }

		private void createNewDetail(string name)
		{
			  try // проверяется, открыт ли компас
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            }
            catch // если не открыт то создаётся инстанс компаса
            {
                kompas = (KompasObject)Activator.CreateInstance(Type.GetTypeFromProgID("KOMPAS.Application.5"));
            }
            if (kompas == null)
                return;
            kompas.Visible = true; // видимость компаса
            ksShaftDoc3d = kompas.Document3D(); // создание 3д документа

            ksShaftDoc3d.Create(false, true); // false - видимый режим, true - деталь
            ksShaftDoc3d = kompas.ActiveDocument3D(); // указатель на интерфейс 3д модели 
            ksShaftDoc3d.author = "Artem";
		}

    }
}
