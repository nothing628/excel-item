using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace IE182
{
    public partial class Main : Form
    {

        private static Random random = new Random();
        private List<List<string>> excelData = new List<List<string>>();

        public Main()
        {
            InitializeComponent();
        }

        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private void listRandomData()
        {
            List<List<string>> test = new List<List<string>>();

            for (int t = 0; t < 30000; t++)
            {
                var childTest = new List<string>();
                for (int c = 0; c < 55; c++)
                {
                    childTest.Add(RandomString(100));
                }
                test.Add(childTest);
            }

            var result = test;

        }

        private void pushExportDataToList(Excel.Range xlRange, int startIndex, int finishIndex)
        {
            int colCount = xlRange.Columns.Count;

            Parallel.For(startIndex, finishIndex, i =>
            {
                List<string> content = new List<string>();
                content.Add(i.ToString());
                Console.WriteLine(i.ToString());
                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        content.Add(xlRange.Cells[i, j].Value2.ToString());
                    else
                        content.Add("");
                }

                excelData.Add(content);
            });
        }

        private void openExcel(string filePath)
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;


            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            int colCount = xlRange.Columns.Count;
            int rowCount = xlRange.Rows.Count;

            int maxRowEachThead = 2500;
            int maxTheads = (rowCount / maxRowEachThead) + 1;
            int remains = rowCount % maxRowEachThead;
            for (int i = 1; i <= maxTheads; i++)
            {
                int startIndex = (i * maxRowEachThead) + 1 - maxRowEachThead;
                int finishIndex = i * maxRowEachThead;
                if (maxTheads == i)
                    finishIndex = rowCount;

                pushExportDataToList(xlRange, startIndex, finishIndex);
            }



            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);

            ////close and release
            //xlWorkbook.Close();
            //Marshal.ReleaseComObject(xlWorkbook);

            ////quit and release
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);

        }

        private void Main_Load(object sender, EventArgs e)
        {

#if DEBUG
            openExcel(@"c:\users\fachri.hawari\desktop\ie182\ie182.xlsx");
            //loadWorkSheet(@"c:\users\fachri.hawari\desktop\ie182\ie182.xlsx");
#endif
        }

        private void buttonChooseFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxFileName.Text = openFileDialog.FileName;
                loadWorkSheet(openFileDialog.FileName);
            }
        }

        private void loadWorkSheet(string path)
        {
            // Load file 
            gridMain.Load(path);

            // Remove unused worksheet
            var willRemoveIndex = gridMain.Worksheets.Count - 1;
            while (willRemoveIndex > 0)
            {
                gridMain.RemoveWorksheet(willRemoveIndex);
                willRemoveIndex--;
            }
            
            var sheet = gridMain.CurrentWorksheet;
            // add space as final
            //sheet.DeleteRows(0, 3);
            //sheet.InsertRows(0, 14);

            //sheet.SetRowsHeight(0, 14, 20);


            // Copy Data from main sheet to final sheet
            // {gridMain.CurrentWorksheet.MaxContentRow}
            var dataRange = $"A1:BC{sheet.MaxContentRow - 10000}";
            var content = sheet.GetPartialGrid(dataRange);
            
            // create new worksheet for the result
            //var finalWorksheet = gridMain.CreateWorksheet("Final");
            //finalWorksheet.SetRows(sheet.MaxContentRow);
            //finalWorksheet.SetCols(sheet.MaxContentCol);

            //gridMain.Worksheets.Add(finalWorksheet);
            //finalWorksheet.SetPartialGrid(dataRange, content);
            
            //// add space as final
            //finalWorksheet.DeleteRows(0, 3);
            //finalWorksheet.InsertRows(0, 14);

            //// Freeze column A-L
            //finalWorksheet.FreezeToCell(0, 12, unvell.ReoGrid.FreezeArea.Left);
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {

                gridMain.Save(saveFileDialog.FileName, unvell.ReoGrid.IO.FileFormat.Excel2007);
            }                
        }

    }
}
