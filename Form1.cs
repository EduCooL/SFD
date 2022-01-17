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

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)//выбор из списка
        {
            
            SelectDocumentation();
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)//документация изменить
        {
            SelectDocumentation();
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)//Журнал октрыть изменить
        {
            this.Hide();
            Form3 form3 = new Form3();
            filename = listBox1.GetItemText(listBox1.SelectedItem);
            if (filename == "")
            {
            filename = listBox1.Items[0].ToString();
            }
            excel.Open(filename);
            excel.Calculated();
            form3.filename = filename;
            string name = "Журнал выполненных работ";
            form3.dataGridView1.DataSource = excel.GetData(name);
            form3.Text = $"СФД: Журнал выполненных работ - {filename}";
            form3.Show();
            
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)// документация сохранить
        {
            excel.Save();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            excel.Close();
        }

        private void ShowList()//Вывод списка
        {
            string path = "C:\\Users\\Public\\Documents\\Documentation\\";
            listBox1.Items.Clear();
            DirectoryInfo dinfo = new DirectoryInfo(path);
            FileInfo[] files = dinfo.GetFiles("*.xlsx");
            foreach (FileInfo filenames in files)
            {
                listBox1.Items.Add(filenames);
            }
        }
        private void SelectDocumentation()//Открытие нужного файла
        {
            this.Hide();
              
            string filename = listBox1.GetItemText(listBox1.SelectedItem);
            if (filename == "")
            {
                filename = listBox1.Items[0].ToString();
            }
            excel.Open(filename);
            excel.Calculated();
            SmetsWindow.Text = $"СФД: Смета - {filename}";
            ((Label)SmetsWindow.Controls["label1"]).Text = $"Список смет - {filename}";
            
            SmetsWindow.Show();
        }
    }
}
