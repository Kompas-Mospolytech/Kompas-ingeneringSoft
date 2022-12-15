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

namespace Шкант_плоский
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static double set_T = 19.4, set_L = 53;

        private void button1_Click(object sender, EventArgs e)
        {
            double set_T = 19.4; //Задаёт новый размер диаметру
            double set_L = 53;

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); // Подключаемся к нашей детали
            ksDocument3D kompas_document_3D = (ksDocument3D)kompas.ActiveDocument3D(); // Получаем интерфейс активной 3Д детали
            ksPart kPart = kompas_document_3D.GetPart((int)Part_Type.pTop_Part); //  Получаем интерфейс ксПарт, используя константы и т.д.
            ksVariableCollection varcoll = kPart.VariableCollection(); // Получаем массив переменных, которые есть в детали
            varcoll.refresh(); //  Обновляем массив этот массив
            ksVariable T = varcoll.GetByName("T", true, true); // Получаем переменную с нашей детали, которая называет T (Важно чтоб переменная была внешняя)
            T.value = set_T; // Задаём новое значение диаметра 
            ksVariable L = varcoll.GetByName("L", true, true);
            L.value = set_L;
            kPart.RebuildModel(); // Перестраиваем модель с новыми размерами 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            double set_T = 19.4; //Задаёт новый размер диаметру
            double set_L = 53.5;

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); // Подключаемся к нашей детали
            ksDocument3D kompas_document_3D = (ksDocument3D)kompas.ActiveDocument3D(); // Получаем интерфейс активной 3Д детали
            ksPart kPart = kompas_document_3D.GetPart((int)Part_Type.pTop_Part); //  Получаем интерфейс ксПарт, используя константы и т.д.
            ksVariableCollection varcoll = kPart.VariableCollection(); // Получаем массив переменных, которые есть в детали
            varcoll.refresh(); //  Обновляем массив этот массив
            ksVariable T = varcoll.GetByName("T", true, true); // Получаем переменную с нашей детали, которая называет T (Важно чтоб переменная была внешняя)
            T.value = set_T; // Задаём новое значение диаметра 
            ksVariable L = varcoll.GetByName("L", true, true);
            L.value = set_L;
            kPart.RebuildModel(); // Перестраиваем модель с новыми размерами 
        }

        private void button3_Click(object sender, EventArgs e)
        {
            double set_T = 23.5; //Задаёт новый размер диаметру
            double set_L = 55;

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); // Подключаемся к нашей детали
            ksDocument3D kompas_document_3D = (ksDocument3D)kompas.ActiveDocument3D(); // Получаем интерфейс активной 3Д детали
            ksPart kPart = kompas_document_3D.GetPart((int)Part_Type.pTop_Part); //  Получаем интерфейс ксПарт, используя константы и т.д.
            ksVariableCollection varcoll = kPart.VariableCollection(); // Получаем массив переменных, которые есть в детали
            varcoll.refresh(); //  Обновляем массив этот массив
            ksVariable T = varcoll.GetByName("T", true, true); // Получаем переменную с нашей детали, которая называет T (Важно чтоб переменная была внешняя)
            T.value = set_T; // Задаём новое значение диаметра 
            ksVariable L = varcoll.GetByName("L", true, true);
            L.value = set_L;
            kPart.RebuildModel(); // Перестраиваем модель с новыми размерами 
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double set_T = 23.5; //Задаёт новый размер диаметру
            double set_L = 56;

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); // Подключаемся к нашей детали
            ksDocument3D kompas_document_3D = (ksDocument3D)kompas.ActiveDocument3D(); // Получаем интерфейс активной 3Д детали
            ksPart kPart = kompas_document_3D.GetPart((int)Part_Type.pTop_Part); //  Получаем интерфейс ксПарт, используя константы и т.д.
            ksVariableCollection varcoll = kPart.VariableCollection(); // Получаем массив переменных, которые есть в детали
            varcoll.refresh(); //  Обновляем массив этот массив
            ksVariable T = varcoll.GetByName("T", true, true); // Получаем переменную с нашей детали, которая называет T (Важно чтоб переменная была внешняя)
            T.value = set_T; // Задаём новое значение диаметра 
            ksVariable L = varcoll.GetByName("L", true, true);
            L.value = set_L;
            kPart.RebuildModel(); // Перестраиваем модель с новыми размерами 
        }

        private void button5_Click(object sender, EventArgs e)
        {
            double set_T = 30; //Задаёт новый размер диаметру
            double set_L = 55;

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); // Подключаемся к нашей детали
            ksDocument3D kompas_document_3D = (ksDocument3D)kompas.ActiveDocument3D(); // Получаем интерфейс активной 3Д детали
            ksPart kPart = kompas_document_3D.GetPart((int)Part_Type.pTop_Part); //  Получаем интерфейс ксПарт, используя константы и т.д.
            ksVariableCollection varcoll = kPart.VariableCollection(); // Получаем массив переменных, которые есть в детали
            varcoll.refresh(); //  Обновляем массив этот массив
            ksVariable T = varcoll.GetByName("T", true, true); // Получаем переменную с нашей детали, которая называет T (Важно чтоб переменная была внешняя)
            T.value = set_T; // Задаём новое значение диаметра 
            ksVariable L = varcoll.GetByName("L", true, true);
            L.value = set_L;
            kPart.RebuildModel(); // Перестраиваем модель с новыми размерами 
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            set_T = System.Convert.ToDouble(comboBox1.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5"); // Подключаемся к нашей детали
            ksDocument3D kompas_document_3D = (ksDocument3D)kompas.ActiveDocument3D(); // Получаем интерфейс активной 3Д детали
            ksPart kPart = kompas_document_3D.GetPart((int)Part_Type.pTop_Part); //  Получаем интерфейс ксПарт, используя константы и т.д.
            ksVariableCollection varcoll = kPart.VariableCollection(); // Получаем массив переменных, которые есть в детали
            varcoll.refresh(); //  Обновляем массив этот массив
            ksVariable T = varcoll.GetByName("T", true, true); // Получаем переменную с нашей детали, которая называет T (Важно чтоб переменная была внешняя)
            T.value = set_T; // Задаём новое значение диаметра 
            ksVariable L = varcoll.GetByName("L", true, true);
            L.value = set_L;
            kPart.RebuildModel(); // Перестраиваем модель с новыми размерами 
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            set_L = System.Convert.ToDouble(comboBox2.Text);
        }
    }
}
