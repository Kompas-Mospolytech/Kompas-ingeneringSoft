using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using System.Runtime.InteropServices;
using KompasAPI7;

namespace KompasNet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            KompasObject kompas;
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
            kompas.Visible = true; // отображение компаса 
            ksDocument3D idoc3d = kompas.Document3D(); // создание детали 3D и документа       
            idoc3d.Create(false, true);
            idoc3d = kompas.ActiveDocument3D();
            
            
            ksPart part = idoc3d.GetPart((int)Part_Type.pTop_Part); //находит верхнюю деталь 
            ksEntity iSketchEntity = part.NewEntity((int)Obj3dType.o3d_sketch); //создание 2д скетча
            SketchDefinition sketchDefinition = iSketchEntity.GetDefinition(); //дать определение скетча, определение хранит настройки
            sketchDefinition.SetPlane(part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY)); //выбор рабочей поверхности 
            iSketchEntity.Create(); //запускаем создание скетча
            ksDocument2D Sketch2d = (ksDocument2D)sketchDefinition.BeginEdit(); //начать редактирование скетча
            Sketch2d.ksLineSeg(0, 0, 10, 0, 1); // KsLineSeg - отрезок, (1-X; 2-Y; 3-X; 4-Y; 5-тип линий)
            Sketch2d.ksLineSeg(10, 0, 10, 10, 1);
            Sketch2d.ksLineSeg(10, 10, 0, 10, 1);
            Sketch2d.ksLineSeg(0, 10, 0, 0, 1);
            sketchDefinition.EndEdit(); //заканчивает редактирование эскиза
            ksEntity bossExtr = part.NewEntity((short)Obj3dType.o3d_bossExtrusion); // задаёт сущность для выдавливания bossExtr - выдавливание с добавлением нового объекта
            BossExtrusionDefinition ExtrDef = bossExtr.GetDefinition(); // задаём значения
            ExtrDef.SetSideParam(true /*если true то он выдавливает в одну сторону, если false, то в другую*/, (short)ksEndTypeEnum.etBlind, 200, 0, false); // пишем свойства для выдавлевания, 200 на сколько выдавливаешь, 0 - угол, etBlind - выдавливание на растояние
            ExtrDef.directionType = (short)ksDirectionTypeEnum.dtNormal; //тип направления dtNormal 
            ExtrDef.SetSketch(iSketchEntity); //пишем название скетча
            bossExtr.Create(); // выдавливание 

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Изменение параметров линии
            /*IApplication application = (IApplication)Marshal.GetActiveObject("KOMPAS.Application.7");
            IKompasDocument kompasDocument = application.ActiveDocument;
            IKompasDocument2D document2D = (IKompasDocument2D)kompasDocument; // подключаемся к нашей детали 

            // подключаемся к интерфейсу активного вида 
            IViewsAndLayersManager viewsAndLayersManager = document2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView view = views.ActiveView;

            //Получаем интерфейс из контейнера (В этом контейнере содержится вся 2D геометрия нашего чертежа)
            IDrawingContainer drawingContainer = (IDrawingContainer)view;
            ILineSegments lineSegments = drawingContainer.LineSegments;   //Выберем все отрезки
            ILineSegment lineSegment = lineSegments.LineSegment[0];   //Конкретный взятый отрезок   

            IDrawingObject drawingObject = (IDrawingObject)lineSegment;*/

            //Изменение размеров детали 
            double set_D = 50; //Задаёт новый размер диаметру
            double set_L = 75;

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); // Подключаемся к нашей детали
            ksDocument3D kompas_document_3D = (ksDocument3D)kompas.ActiveDocument3D(); // Получаем интерфейс активной 3Д детали
            ksPart kPart = kompas_document_3D.GetPart((int)Part_Type.pTop_Part); //  Получаем интерфейс ксПарт, используя константы и т.д.
            ksVariableCollection varcoll = kPart.VariableCollection(); // Получаем массив переменных, которые есть в детали
            varcoll.refresh(); //  Обновляем массив этот массив
            ksVariable D = varcoll.GetByName("D", true, true); // Получаем переменную с нашей детали, которая называет D (Важно чтоб переменная была внешняя)
            D.value = set_D; // Задаём новое значение диаметра 
            ksVariable L = varcoll.GetByName("L", true, true); 
            L.value = set_L;
            kPart.RebuildModel(); // Перестраиваем модель с новыми размерами 

        }
    }
}
