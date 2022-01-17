using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Documentation
{
    public partial class Form4 : Form
    {
        ExcelWork excel = new ExcelWork();
        string path = "C:\\Users\\Public\\Documents\\Documentation\\";
        public string filename;
        public string name;
        public Form4()
        {
            InitializeComponent();
        }

        private void сохранитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Save();
        }

}
