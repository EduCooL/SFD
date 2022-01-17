using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace Documentation
{
    public partial class Form2 : Form
    {
        ExcelWork excel = new ExcelWork();
        public string filename;
        public Form2()
        { 
            
            InitializeComponent();
            

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            GetList();
        }

        private void button2_Click(object sender, EventArgs e)//Создать
        {
            
            string name = Microsoft.VisualBasic.Interaction.InputBox("Введите текст:");
            excel.CreateSmeta(name);
            GetList();

        }

        private void button1_Click(object sender, EventArgs e)//Обновить
        {
            GetList();
        }
 
        private void button3_Click(object sender, EventArgs e)//сохранить
        {
            Saved();
        }
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)//Смета сохранить
        {
            Saved();
        }
        private void изменитьToolStripMenuItem1_Click(object sender, EventArgs e)//Смета изменить
        {
            ChangingSmeta();
        }
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)//Документация открыть
        {
            this.Close();
        }
        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
        }
        private void GetList()
        {
            listBox1.Items.Clear();
            var smets = excel.GetSheetsList();
            

            foreach (var name in smets)
            {
                listBox1.Items.Add(name);
            }
        }

        private void Saved()
        {
            excel.Save();
            this.Close();
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ChangingSmeta();
        }
        private void ChangingSmeta()
        {
            this.Hide();
            Form4 form4 = new Form4();
            filename = label1.Text.Remove(0, label1.Text.LastIndexOf(" ") + 1);
            string name = "Журнал выполненных работ";
            name = listBox1.GetItemText(listBox1.SelectedItem);
            if (name == "" || name == null)
            {
                name = listBox1.Items[0].ToString();
            }
            form4.name = name;
            form4.filename = filename.ToString();
            excel.Open(filename);
            excel.Calculated();
            form4.dataGridView1.DataSource = excel.GetData(name);
            form4.Text = $"СФД: Смета - {name}";
            ((Label)form4.Controls["label1"]).Text = $"{name}";
            form4.Show();
        }
    }
}
