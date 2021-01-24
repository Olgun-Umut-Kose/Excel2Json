using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel2Json.Concrate;
using Formtest.Classes;

namespace Formtest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls";
            ofd.RestoreDirectory = true;
            ofd.Title = "Excel Dosyası Seçiniz..";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string yol = ofd.FileName;
                Excel2JsonDal excel2Json = new Excel2JsonDal(yol,new Excel12());
                OleDbDataReader read = excel2Json.Read("ADMİN");
                string json = excel2Json.ConvertJson(read, new Sutun());
                MessageBox.Show(json);
            }
        }
    }
}
