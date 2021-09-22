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

namespace ExcelRepeatFixer
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorkSheet;
        object missingValue = System.Reflection.Missing.Value;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowseFiles_Click(object sender, EventArgs e)
        {
            DialogResult result = openExcelFile.ShowDialog();
            if (result = DialogResult.OK)
            {
                string file = openExcelFile.FileName;
                xlApp = new Excel.Application();
                // Didn't follow the guid exactly here, I don't want to type all those arguments, hopefully it doesn't cause bugs
                xlWorkbook = xlApp.Workbooks.Open(file);
                xlWorkSheet = (Excel.Worksheet) xlWorkbook.Worksheets.Get_Item(1); // I assume this gets the first sheet, user may want selector

            }
        }
    }
}
