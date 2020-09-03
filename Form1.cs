using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using TestML;

namespace WinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private OpenFileDialog openFileDialog1;
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            comboBox1.Text = openFileDialog1.FileName;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        // Создание массивов для прогнозируемых значений

        float[] arrPredict1 = new float[4]; 
        float[] arrPredict2 = new float[4];
        float[] arrPredict3 = new float[4];
        float[] arrPredict4 = new float[4];

        // Создание массивов для значений групп коэффициентов

        float[] arrRentab = new float[32];
        float[] arrLikvid = new float[32];
        float[] arrDelActiv = new float[32];
        float[] arrFinUst = new float[32];

        private void button2_Click(object sender, EventArgs e)
        {
            ProgramML.PrepareExcel(openFileDialog1.FileName, ref arrRentab, ref arrLikvid, ref arrDelActiv, ref arrFinUst);
            ProgramML.BuildPredictModel(ref arrPredict1,ref arrPredict2, ref arrPredict3, ref arrPredict4);
            //MessageBox.Show(test.ToString());
            //MessageBox.Show(arrPredict1[0].ToString()+ " " + arrPredict1[1].ToString() + " " + arrPredict1[2].ToString() + " " + arrPredict1[3].ToString() );

            //for (int i = 0; i < 32; i++)
            //    arrRentab[i] = i;

            //for (int i = 0; i < 32; i++)
            //    arrLikvid[i] = i;

            //for (int i = 0; i < 32; i++)
            //    arrDelActiv[i] = i;

            //for (int i = 0; i < 32; i++)
            //    arrFinUst[i] = i;

            BuildGraphic a = new BuildGraphic(ref arrRentab, ref arrLikvid, ref arrDelActiv, ref arrFinUst, ref arrPredict1, ref arrPredict2, 
                ref arrPredict3, ref arrPredict4);
            a.ShowDialog();

        }
    }
}
