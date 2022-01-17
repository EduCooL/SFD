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

       
    }
}
