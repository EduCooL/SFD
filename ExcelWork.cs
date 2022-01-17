using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using System.Windows.Forms;

namespace Documentation
{
    
    public class ExcelWork
    {
        static Excel.Application ObjWorkExcel;
        static Excel.Workbook ObjWorkBook;
        static Excel.Worksheet Sheet;
        static Excel.Workbooks Workbooks;
        string path = "C:\\Users\\Public\\Documents\\Documentation\\";
        

       

        public ExcelWork()
        {
            
            ObjWorkExcel = new Excel.Application();
            ObjWorkBook = ObjWorkExcel.Workbooks.Add(Type.Missing);
            Sheet = (Excel.Worksheet)ObjWorkBook.ActiveSheet;
            Workbooks = ObjWorkExcel.Workbooks;

        }
        public void Calculated()
        {
            for (var i = 1; i <= ObjWorkBook.Sheets.Count; i++)
            {
                Sheet = ObjWorkBook.Sheets[i];
                var range = Sheet.UsedRange;
                int a, b, p, t;
                string s = "";
                for(var j = 2;j <= range.Rows.Count; j++ )
                if((range.Cells[j, "C"] as Excel.Range).Value2 != null &&
                        (range.Cells[j, "D"] as Excel.Range).Value2 != null &&
                        int.TryParse(range.Cells[j, "C"].Text, out t) &&
                        int.TryParse(range.Cells[j, "D"].Text, out t))
                {
                     a = int.Parse(Sheet.Cells[j, "C"].Text);
                     b = int.Parse(Sheet.Cells[j, "D"].Text);
                     p = a * b;
                     s = p.ToString();
                     Set(column: "E", row: j, s);
                }
            }
            Save();
        }
        public void Create(string name)
        {

            try
            {
                if (name != "")
                    ObjWorkBook.SaveAs(path + $"{name}.xlsx",
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Excel.XlSaveAsAccessMode.xlShared,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing);

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            }
            finally
            {
                Marshal.ReleaseComObject(ObjWorkBook);
            }
        }
        public void CreateWorkJournal()
        {
            
            try
            {
                Sheet.Name = "Журнал выполненных работ";
                Set(column: "A", row: 1, "#");
                Set(column: "B", row: 1, "Наименование работы");
                Set(column: "C", row: 1, "Объем работы");
                Set(column: "D", row: 1, "Цена за единицу объема, руб");
                Set(column: "E", row: 1, "Итоговая цена");
            }
            finally
            {
                Marshal.ReleaseComObject(Sheet);
            }
        }
        public void CreateSmeta(string name)
        {
            var sheets = ObjWorkBook.Sheets;
            Sheet = sheets.Add();
            Sheet.Name = name;
            Set(column: "A", row: 1, "#");
            Set(column: "B", row: 1, "Наименование");
            Set(column: "C", row: 1, "Объем");
            Set(column: "D", row: 1, "Цена за единицу объема, руб");
            Set(column: "E", row: 1, "Итоговая цена");

        }
        private void Set(string column, int row, string data)
        {
            Excel.Range range = Sheet.Cells;
            try
            {
                range[row, column] = data;
                range[row, column].EntireColumn.AutoFit();
            }
            finally
            {
                Marshal.ReleaseComObject(range);
            }
        }
       
        public void Open(string filename)
        {
            ObjWorkBook =  Workbooks.Open(path + filename);
        }
        public List<string> GetSheetsList()
        {
            
            var sheets = ObjWorkBook.Sheets;
            var names = new List<string>();
            for(var i = 1; i <= sheets.Count; i++)
                if(sheets[i].Name != "Журнал выполненных работ")
                names.Add(sheets[i].Name);
            Marshal.ReleaseComObject(sheets);
            return names;
        }
        public void Save()
        {
            
            ObjWorkBook.Save();

        }

        public void Close()
        {
            try
            {
                ObjWorkExcel.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(ObjWorkExcel);
            }

            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }



    }
}
