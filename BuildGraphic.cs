using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ZedGraph;

namespace WinForm
{
    public partial class BuildGraphic : Form
    {
        public BuildGraphic()
        {
            InitializeComponent();

        }

        public BuildGraphic(ref float[] arr1, ref float[] arr2, ref float[] arr3, ref float[] arr4,
            ref float[] arrPred1, ref float[] arrPred2, ref float[] arrPred3, ref float[] arrPred4)
        {
            InitializeComponent();

            //------------------------------------------------------------------------------------------------------------------

            // Первый график

            ZedGraphControl zedGraph1 = new ZedGraphControl();

            zedGraph1.Location = new System.Drawing.Point(20, 20);
            zedGraph1.Name = "zedGraph";
            zedGraph1.Size = new System.Drawing.Size(500, 300);
            this.Controls.Add(zedGraph1);

            GraphPane myPane = zedGraph1.GraphPane;

            // Работа с осями и надписями графика

            myPane.Title.Text = "Показатели рентабельности";
            myPane.XAxis.Title.Text = "Число кварталов";
            myPane.YAxis.Title.Text = "Значения данной группы коэффициентов";

            // Добавление информации на график

            PointPairList list = new PointPairList();
            for (int i = 0; i < 32; i++)
            {
                list.Add((double)(i+1), (double)arr1[i]);
            }
            int j1 = 0;
            for (int i = 33; i < 37; i++)
                list.Add((double)i, arrPred1[j1++]);

            // Работа с цветом

            LineItem myCurve = myPane.AddCurve("f(x)",
               list, Color.Green, SymbolType.Diamond);

            // Установка точки пересечения осей Х и Y

            myPane.YAxis.Cross = 0.0;

            // Настройка свойств отображения осей

            myPane.Chart.Border.IsVisible = false;
            myPane.XAxis.MajorTic.IsOpposite = false;
            myPane.XAxis.MinorTic.IsOpposite = false;
            myPane.YAxis.MajorTic.IsOpposite = false;
            myPane.YAxis.MinorTic.IsOpposite = false;

            // Внесение изменений в график

            zedGraph1.AxisChange();
            zedGraph1.Refresh();

            //------------------------------------------------------------------------------------------------------------------

            // Второй график

            ZedGraphControl zedGraph2 = new ZedGraphControl();

            zedGraph2.Location = new System.Drawing.Point(540, 20);
            zedGraph2.Name = "zedGraph";
            zedGraph2.Size = new System.Drawing.Size(500, 300);
            this.Controls.Add(zedGraph2);

            GraphPane myPane2 = zedGraph2.GraphPane;
            
            myPane2.Title.Text = "Показатели ликвидности";
            myPane2.XAxis.Title.Text = "Число кварталов";
            myPane2.YAxis.Title.Text = "Значения данной группы коэффициентов";

            PointPairList list2 = new PointPairList();
            for (int i = 0; i < 32; i++)
            {
                list2.Add((double)(i + 1), (double)arr2[i]);
            }
            int j2 = 0;
            for (int i = 33; i < 37; i++)
                list2.Add((double)i, arrPred2[j2++]);

            LineItem myCurve2 = myPane2.AddCurve("f(x)",
               list2, Color.Green, SymbolType.Diamond);

            myPane2.YAxis.Cross = 0.0;

            myPane2.Chart.Border.IsVisible = false;
            myPane2.XAxis.MajorTic.IsOpposite = false;
            myPane2.XAxis.MinorTic.IsOpposite = false;
            myPane2.YAxis.MajorTic.IsOpposite = false;
            myPane2.YAxis.MinorTic.IsOpposite = false;

            zedGraph2.AxisChange();
            zedGraph2.Refresh();

            //------------------------------------------------------------------------------------------------------------------

            // Третий график

            ZedGraphControl zedGraph3 = new ZedGraphControl();

            zedGraph3.Location = new System.Drawing.Point(20, 340);
            zedGraph3.Name = "zedGraph";
            zedGraph3.Size = new System.Drawing.Size(500, 300);
            this.Controls.Add(zedGraph3);

            GraphPane myPane3 = zedGraph3.GraphPane;

            myPane3.Title.Text = "Показатели деловой активности";
            myPane3.XAxis.Title.Text = "Число кварталов";
            myPane3.YAxis.Title.Text = "Значения данной группы коэффициентов";

            PointPairList list3 = new PointPairList();
            for (int i = 0; i < 32; i++)
            {
                list3.Add((double)(i + 1), (double)arr3[i]);
            }
            int j3 = 0;
            for (int i = 33; i < 37; i++)
                list3.Add((double)i, arrPred3[j3++]);

            LineItem myCurve3 = myPane3.AddCurve("f(x)",
               list3, Color.Green, SymbolType.Diamond);

            myPane3.YAxis.Cross = 0.0;

            myPane3.Chart.Border.IsVisible = false;
            myPane3.XAxis.MajorTic.IsOpposite = false;
            myPane3.XAxis.MinorTic.IsOpposite = false;
            myPane3.YAxis.MajorTic.IsOpposite = false;
            myPane3.YAxis.MinorTic.IsOpposite = false;

            zedGraph3.AxisChange();
            zedGraph3.Refresh();

            //------------------------------------------------------------------------------------------------------------------

            // Четвертый график

            ZedGraphControl zedGraph4 = new ZedGraphControl();

            zedGraph4.Location = new System.Drawing.Point(540, 340);
            zedGraph4.Name = "zedGraph";
            zedGraph4.Size = new System.Drawing.Size(500, 300);
            this.Controls.Add(zedGraph4);

            GraphPane myPane4 = zedGraph4.GraphPane;

            myPane4.Title.Text = "Показатели финансовой устойчивости";
            myPane4.XAxis.Title.Text = "Число кварталов";
            myPane4.YAxis.Title.Text = "Значения данной группы коэффициентов";

            PointPairList list4 = new PointPairList();
            for (int i = 0; i < 32; i++)
            {
                list4.Add((double)(i + 1), (double)arr4[i]);
            }
            int j4 = 0;
            for (int i = 33; i < 37; i++)
                list4.Add((double)i, arrPred4[j4++]);

            LineItem myCurve4 = myPane4.AddCurve("f(x)",
               list4, Color.Green, SymbolType.Diamond);

            myPane4.YAxis.Cross = 0.0;

            myPane4.Chart.Border.IsVisible = false;
            myPane4.XAxis.MajorTic.IsOpposite = false;
            myPane4.XAxis.MinorTic.IsOpposite = false;
            myPane4.YAxis.MajorTic.IsOpposite = false;
            myPane4.YAxis.MinorTic.IsOpposite = false;

            zedGraph4.AxisChange();
            zedGraph4.Refresh();
        }      
    }
}
