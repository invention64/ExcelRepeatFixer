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
using System.IO;

namespace ExcelRepeatFixer
{
    public partial class RepeatFixer : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorkSheet;
        object missingValue = System.Reflection.Missing.Value;
        string excelFilename = null;

        public RepeatFixer()
        {
            InitializeComponent();
        }

        private void btnBrowseFiles_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Showing Dialog");
            DialogResult result = openExcelFile.ShowDialog();
            Console.WriteLine("Dialog Ended");
        }

        private void openExcelFile_FileOk(object sender, CancelEventArgs e)
        {
            excelFilename = openExcelFile.FileName;

            fileSelectedTxt.Text = String.Copy(excelFilename).Remove(0,excelFilename.LastIndexOf("\\") + 1);
            Console.WriteLine("Dialog OK");

            xlApp = new Excel.Application();
            // Didn't follow the guide exactly here, I don't want to type all those arguments, hopefully it doesn't cause bugs. It did!
            xlWorkbook = xlApp.Workbooks.Open(excelFilename, 0, false,Excel.XlPlatform.xlWindows, "\t", false, true);
            xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.Item[1]; // I assume this gets the first sheet, user may want selector
            Console.WriteLine("Excel Sheet Loaded");

            // check how many cells are filled on the top row and fill a list to be displayed in the drop down
            List<string> columnNames = new List<string>();
            Excel.Range range = null; // these ranges are strange
            int xIndex = 1; //Excel is INDEXED BY 1!
            while (xIndex == 1 || range.Value != null)
            { 
                // get a range? for each column
                range = xlWorkSheet.Cells[1, xIndex] as Excel.Range;
                // get the text from each column
                if (range.Value != null) {
                    columnNames.Add(range.Value.ToString());
                }
                xIndex++;
            }
            Console.WriteLine("Column names read");

            // set the fieldComboBox
            fieldCombo.Items.AddRange(columnNames.ToArray());
            fieldCombo.SelectedIndex = 0; // this will set the combobox to have text in it
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (excelFilename == null)
            {
                lblError.Text = "Must select file";
                lblError.Visible = true;
                return;
            }
            lblError.Visible = false;
            Console.WriteLine("Start Excel Process");
            // check if user wants to backup file
            if (backupChk.Checked)
            {
                // get the path
                string path = excelFilename.Substring(0, excelFilename.LastIndexOf("\\"));

                // copy the file
                File.Copy(excelFilename, path + "\\backup-" + fileSelectedTxt.Text);
            }
        }
    }
}
