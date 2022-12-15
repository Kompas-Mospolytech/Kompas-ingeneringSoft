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
using System.IO;


namespace Шкант
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static double set_D = 5, set_L = 25;
        public static string set_Material;

        private void button1_Click(object sender, EventArgs e)
        {
            double set_D = 5; //Задаёт новый размер диаметру
            double set_L = 25;

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

        private void button2_Click(object sender, EventArgs e)
        {
            double set_D = 6; //Задаёт новый размер диаметру
            double set_L = 30;

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

        private void button3_Click(object sender, EventArgs e)
        {
            double set_D = 6; //Задаёт новый размер диаметру
            double set_L = 45;

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

        private void button4_Click(object sender, EventArgs e)
        {
            double set_D = 8; //Задаёт новый размер диаметру
            double set_L = 30;

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

        private void button5_Click(object sender, EventArgs e)
        {
            double set_D = 8; //Задаёт новый размер диаметру
            double set_L = 35;

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

        private void button6_Click(object sender, EventArgs e)
        {
            double set_D = 8; //Задаёт новый размер диаметру
            double set_L = 40;

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

        private void button7_Click(object sender, EventArgs e)
        {
            double set_D = 8; //Задаёт новый размер диаметру
            double set_L = 50;

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

        private void button8_Click(object sender, EventArgs e)
        {
            double set_D = 8; //Задаёт новый размер диаметру
            double set_L = 60;

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

        private void button9_Click(object sender, EventArgs e)
        {
            double set_D = 10; //Задаёт новый размер диаметру
            double set_L = 45;

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

        private void button10_Click(object sender, EventArgs e)
        {
           

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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            set_L = System.Convert.ToDouble(comboBox2.Text);
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            set_D = System.Convert.ToDouble(comboBox1.Text);
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            set_Material = System.Convert.ToString(comboBox3.Text);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            IApplication application = (IApplication)Marshal.GetActiveObject("KOMPAS.Application.7");
            IKompasDocument3D document3d = (IKompasDocument3D)application.ActiveDocument;

            IPart7 part7 = document3d.TopPart;

            part7.SetMaterial(set_Material, 7.85);

            part7.Update();
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            set_L = System.Convert.ToDouble(comboBox2.Text);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            set_D = System.Convert.ToDouble(comboBox1.Text);
        }
    }
    }
   
