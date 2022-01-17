using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace Documentation
{
    public partial class Form3 : Form
    {
        ExcelWork excel = new ExcelWork();
        string path = "C:\\Users\\Public\\Documents\\Documentation\\";
        public string filename;
        public Form3()
        {
            InitializeComponent();
            
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 form1 = new Form1(); 
            form1.Show();
        }

        private void сохранитьToolStripMenuItem2_Click(object sender, EventArgs e)//Журнал сохранить
        {
            Save();
        }
        private void Save()
        {
            Excel.Application exApp = new Excel.Application();
            var book = exApp.Workbooks.Open(path + filename);
            var count = book.Sheets.Count;
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.Sheets[count];
            int i, j;
            for (i = 0; i <= dataGridView1.RowCount - 2; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    wsh.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                    object p = wsh.Cells[i + 2, j + 1].EntireColumn.AutoFit();
                }
            }
            book.Save();
            this.Close();
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
