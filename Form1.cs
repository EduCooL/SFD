using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static Documentation.ExcelWork;

namespace Documentation
{
    public partial class Form1 : Form
    {
        Form2 SmetsWindow;
        ExcelWork excel = new ExcelWork();
        public string filename = "";
        public Form1()
        {
            InitializeComponent();
            string root = @"C:\Users\Public\Documents";
            string subDir = @"Documentation";
            DirectoryInfo directory = new DirectoryInfo(root);
            if (!directory.Exists)
            {
                directory.Create();
            }
            DirectoryInfo newDir = directory.CreateSubdirectory(subDir);
            SmetsWindow = new Form2();
            ShowList();

        }

        private void button2_Click(object sender, EventArgs e)//Создать
        { 
            excel.CreateWorkJournal();
            string name = Microsoft.VisualBasic.Interaction.InputBox("Введите текст:");
            excel.Create(name);
            ShowList();

        }
        private void button1_Click(object sender, EventArgs e)//Обновить
        {
            ShowList();
        }

        private void button3_Click(object sender, EventArgs e)//Экспорт
        {
        
            
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)//смета открыть
        {
            SelectDocumentation();
            SmetsWindow.Show();
            this.Hide();       
        }

        
    }
}
