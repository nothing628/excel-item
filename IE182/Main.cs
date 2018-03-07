using LiteDB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using unvell.ReoGrid;

namespace IE182
{
    public partial class Main : Form
    {

        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
        }

        private void buttonChooseFile_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBoxFileName.Text = openFileDialog.FileName;
                    loadWorkSheet(openFileDialog.FileName);
                    buttonStoreToDB.Enabled = true;
                }
            }
        }

        private void storeToDB() {
            var currentSheet = gridMain.CurrentWorksheet;
            var maxRow = currentSheet.RowCount;
            var maxCol = currentSheet.ColumnCount;

            if (File.Exists(@"IE182.db"))
                File.Delete(@"IE182.db");

            // Open database (or create if not exits)
            using (var db = new LiteDatabase(@"IE182.db"))
            {
                // Get customer collection
                var items = db.GetCollection<Item>("items");
                
                for (int i = 5; i <= maxRow; i++)
                {
                    var firstCell = currentSheet[$"L{i}"];

                    if (firstCell == null || firstCell.ToString().Length == 0) //This empty data, ignore it
                        break;

                    var item = new Item
                    {
                        A = currentSheet[$"A{i}"] != null ? currentSheet[$"A{i}"].ToString() : "",
                        B = currentSheet[$"B{i}"] != null ? currentSheet[$"B{i}"].ToString() : "",
                        C = currentSheet[$"C{i}"] != null ? currentSheet[$"C{i}"].ToString() : "",
                        D = currentSheet[$"D{i}"] != null ? currentSheet[$"D{i}"].ToString() : "",
                        E = currentSheet[$"E{i}"] != null ? currentSheet[$"E{i}"].ToString() : "",
                        F = currentSheet[$"F{i}"] != null ? currentSheet[$"F{i}"].ToString() : "",
                        G = currentSheet[$"G{i}"] != null ? currentSheet[$"G{i}"].ToString() : "",
                        H = currentSheet[$"H{i}"] != null ? currentSheet[$"H{i}"].ToString() : "",
                        I = currentSheet[$"I{i}"] != null ? currentSheet[$"I{i}"].ToString() : "",
                        J = currentSheet[$"J{i}"] != null ? currentSheet[$"J{i}"].ToString() : "",
                        K = currentSheet[$"K{i}"] != null ? currentSheet[$"K{i}"].ToString() : "",
                        L = currentSheet[$"L{i}"] != null ? currentSheet[$"L{i}"].ToString() : "",
                        M = currentSheet[$"M{i}"] != null ? currentSheet[$"M{i}"].ToString() : "",
                        N = currentSheet[$"N{i}"] != null ? currentSheet[$"N{i}"].ToString() : "",
                        O = currentSheet[$"O{i}"] != null ? currentSheet[$"O{i}"].ToString() : "",
                        P = currentSheet[$"P{i}"] != null ? currentSheet[$"P{i}"].ToString() : "",
                        Q = currentSheet[$"Q{i}"] != null ? currentSheet[$"Q{i}"].ToString() : "",
                        R = currentSheet[$"R{i}"] != null ? currentSheet[$"R{i}"].ToString() : "",
                        S = currentSheet[$"S{i}"] != null ? currentSheet[$"S{i}"].ToString() : "",
                        T = currentSheet[$"T{i}"] != null ? currentSheet[$"T{i}"].ToString() : "",
                        U = currentSheet[$"U{i}"] != null ? currentSheet[$"U{i}"].ToString() : "",
                        V = currentSheet[$"V{i}"] != null ? currentSheet[$"V{i}"].ToString() : "",
                        W = currentSheet[$"W{i}"] != null ? currentSheet[$"W{i}"].ToString() : "",
                        X = currentSheet[$"X{i}"] != null ? currentSheet[$"X{i}"].ToString() : "",
                        Y = currentSheet[$"Y{i}"] != null ? currentSheet[$"Y{i}"].ToString() : "",
                        Z = currentSheet[$"Z{i}"] != null ? currentSheet[$"Z{i}"].ToString() : "",
                        AA = currentSheet[$"AA{i}"] != null ? currentSheet[$"AA{i}"].ToString() : "",
                        AB = currentSheet[$"AB{i}"] != null ? currentSheet[$"AB{i}"].ToString() : "",
                        AC = currentSheet[$"AC{i}"] != null ? currentSheet[$"AC{i}"].ToString() : "",
                        AD = currentSheet[$"AD{i}"] != null ? currentSheet[$"AD{i}"].ToString() : "",
                        AE = currentSheet[$"AE{i}"] != null ? currentSheet[$"AE{i}"].ToString() : "",
                        AF = currentSheet[$"AF{i}"] != null ? currentSheet[$"AF{i}"].ToString() : "",
                        AG = currentSheet[$"AG{i}"] != null ? currentSheet[$"AG{i}"].ToString() : "",
                        AH = currentSheet[$"AH{i}"] != null ? currentSheet[$"AH{i}"].ToString() : "",
                        AI = currentSheet[$"AI{i}"] != null ? currentSheet[$"AI{i}"].ToString() : "",
                        AJ = currentSheet[$"AJ{i}"] != null ? currentSheet[$"AJ{i}"].ToString() : "",
                        AK = currentSheet[$"AK{i}"] != null ? currentSheet[$"AK{i}"].ToString() : "",
                        AL = currentSheet[$"AL{i}"] != null ? currentSheet[$"AL{i}"].ToString() : "",
                        AM = currentSheet[$"AM{i}"] != null ? currentSheet[$"AM{i}"].ToString() : "",
                        AN = currentSheet[$"AN{i}"] != null ? currentSheet[$"AN{i}"].ToString() : "",
                        AO = currentSheet[$"AO{i}"] != null ? currentSheet[$"AO{i}"].ToString() : "",
                        AP = currentSheet[$"AP{i}"] != null ? currentSheet[$"AP{i}"].ToString() : "",
                        AQ = currentSheet[$"AQ{i}"] != null ? currentSheet[$"AQ{i}"].ToString() : "",
                        AR = currentSheet[$"AR{i}"] != null ? currentSheet[$"AR{i}"].ToString() : "",
                        AS = currentSheet[$"AS{i}"] != null ? currentSheet[$"AS{i}"].ToString() : "",
                        AT = currentSheet[$"AT{i}"] != null ? currentSheet[$"AT{i}"].ToString() : "",
                        AU = currentSheet[$"AU{i}"] != null ? currentSheet[$"AU{i}"].ToString() : "",
                        AV = currentSheet[$"AV{i}"] != null ? currentSheet[$"AV{i}"].ToString() : "",
                        AW = currentSheet[$"AW{i}"] != null ? currentSheet[$"AW{i}"].ToString() : "",
                        AX = currentSheet[$"AX{i}"] != null ? currentSheet[$"AX{i}"].ToString() : "",
                        AY = currentSheet[$"AY{i}"] != null ? currentSheet[$"AY{i}"].ToString() : "",
                        AZ = currentSheet[$"AZ{i}"] != null ? currentSheet[$"AZ{i}"].ToString() : "",
                        BA = currentSheet[$"BA{i}"] != null ? currentSheet[$"BA{i}"].ToString() : "",
                        BB = currentSheet[$"BB{i}"] != null ? currentSheet[$"BB{i}"].ToString() : "",
                        BC = currentSheet[$"BC{i}"] != null ? currentSheet[$"BC{i}"].ToString() : ""
                    };

                    items.Insert(item);
                }
                
                // Index document using document Name property
                items.EnsureIndex(x => x.L);
                items.EnsureIndex(x => x.M);
                items.EnsureIndex(x => x.AF);

            }

            GC.Collect();
        }

        private void loadWorkSheet(string path)
        {
            try
            {
                // Load file 
                gridMain.Load(path, unvell.ReoGrid.IO.FileFormat.Excel2007, Encoding.Unicode);
                
                // Remove unused worksheet
                var willRemoveIndex = gridMain.Worksheets.Count - 1;
                while (willRemoveIndex > 0)
                {
                    gridMain.RemoveWorksheet(willRemoveIndex);
                    willRemoveIndex--;
                }
                
                gridMain.CurrentWorksheet.SetRowsHeight(3, 1, 20);
            } catch (Exception ex)
            {
                // Failed to load
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            GC.Collect();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    gridMain.Save(saveFileDialog.FileName, unvell.ReoGrid.IO.FileFormat.Excel2007);
            }        
        }

        private void buttonStoreToDB_Click(object sender, EventArgs e)
        {
            storeToDB();
            ProcessDB();
            buttonProcess.Enabled = true;
        }

        private void buttonProcess_Click(object sender, EventArgs e)
        {
            ProcessDB();
        }

        private void ProcessDB()
        {
            //Check file Database
            if (!File.Exists("IE182.db"))
                return;

            //Open Database
            using (var db = new LiteDatabase("IE182.db"))
            {
                var workshit = gridMain.CreateWorksheet("Final");
                //Get Collections
                var collect = db.GetCollection<Item>("items");
                var items = collect.FindAll();
                var grp_items = (from a in items
                                 group a by new
                                 {
                                     MatlNo = a.AF,
                                     MatlNa = a.AG,
                                     Unit = a.AH,
                                     Supplier = a.AL + a.AM,
                                     MatlLT = a.AO,
                                     TransDat = a.AP,
                                 } into b
                                 select new
                                 {
                                     b.Key.MatlNo,
                                     b.Key.MatlNa,
                                     b.Key.Unit,
                                     b.Key.Supplier,
                                     b.Key.MatlLT,
                                     b.Key.TransDat,
                                     POs = b.ToList()
                                 }).ToList();

                var row_pos = 1;
                var col_pos = 22; // Column 'W'
                var row_item = 14;
                foreach (var item in grp_items)
                {
                    // Group print
                    AutoAppend(workshit, row_pos + 12, col_pos + 15);

                    workshit[row_pos + 0, col_pos] = "料號";
                    workshit[row_pos + 1, col_pos] = "料名";
                    workshit[row_pos + 2, col_pos] = "規格";
                    workshit[row_pos + 3, col_pos] = "單位";
                    workshit[row_pos + 4, col_pos] = "供應商";
                    workshit[row_pos + 5, col_pos] = "MOQ";
                    workshit[row_pos + 6, col_pos] = "交期天數";
                    workshit[row_pos + 7, col_pos] = "運輸天數";
                    workshit[row_pos + 8, col_pos] = "提前加工天數";
                    workshit[row_pos + 9, col_pos] = "安心庫存天數";
                    workshit[row_pos + 10, col_pos] = "下單天數";

                    workshit[row_pos + 0, col_pos + 1] = item.MatlNo;
                    workshit[row_pos + 1, col_pos + 1] = item.MatlNa;
                    workshit[row_pos + 3, col_pos + 1] = item.Unit;
                    workshit[row_pos + 4, col_pos + 1] = item.Supplier;
                    workshit[row_pos + 6, col_pos + 1] = item.MatlLT;
                    workshit[row_pos + 7, col_pos + 1] = item.TransDat;

                    workshit[row_pos + 12, col_pos + 0] = "單位用量";
                    workshit[row_pos + 12, col_pos + 1] = "損耗率";
                    workshit[row_pos + 12, col_pos + 2] = "需求量";
                    workshit[row_pos + 12, col_pos + 3] = "已採購數量";
                    workshit[row_pos + 12, col_pos + 4] = "採購單號";
                    workshit[row_pos + 12, col_pos + 5] = "採購日期";
                    workshit[row_pos + 12, col_pos + 6] = "已發出數量";
                    workshit[row_pos + 12, col_pos + 7] = "領料單號";
                    workshit[row_pos + 12, col_pos + 8] = "領料日期";
                    workshit[row_pos + 12, col_pos + 9] = "採購數量(未交量)";
                    workshit[row_pos + 12, col_pos + 10] = "訂單號碼(採購單號)";
                    workshit[row_pos + 12, col_pos + 11] = "入庫日期";
                    workshit[row_pos + 12, col_pos + 12] = "數量";
                    workshit[row_pos + 12, col_pos + 13] = "單號";
                    workshit[row_pos + 12, col_pos + 14] = "出庫日期";
                    workshit[row_pos + 12, col_pos + 15] = "期末庫存";

                    foreach (var po in item.POs)
                    {
                        // Write PO
                        workshit[row_item, 0] = po.A;
                        workshit[row_item, 1] = "'" + po.B;
                        workshit[row_item, 2] = po.C;
                        workshit[row_item, 3] = "'" + po.D;
                        workshit[row_item, 4] = po.E;
                        workshit[row_item, 5] = po.F;
                        workshit[row_item, 6] = po.G;
                        workshit[row_item, 7] = po.H;
                        workshit[row_item, 8] = po.I;
                        workshit[row_item, 9] = po.J;
                        workshit[row_item, 10] = po.K;
                        workshit[row_item, 11] = "'" + po.L;
                        workshit[row_item, 12] = po.M;
                        workshit[row_item, 13] = "'" + po.N;
                        workshit[row_item, 14] = po.O;
                        workshit[row_item, 15] = "'" + po.P;
                        workshit[row_item, 16] = "'" + po.Q;
                        workshit[row_item, 17] = "'" + po.R;
                        workshit[row_item, 18] = po.S;
                        workshit[row_item, 19] = po.T;
                        workshit[row_item, 20] = po.U;
                        workshit[row_item, 21] = po.V;

                        workshit[row_item, col_pos + 0] = po.AI;
                        workshit[row_item, col_pos + 1] = po.AJ;
                        workshit[row_item, col_pos + 2] = po.AK;
                        workshit[row_item, col_pos + 3] = po.AQ;
                        workshit[row_item, col_pos + 4] = po.AR;
                        workshit[row_item, col_pos + 5] = po.AS;
                        workshit[row_item, col_pos + 6] = po.AT;
                        workshit[row_item, col_pos + 7] = po.AU;
                        workshit[row_item, col_pos + 8] = po.AV;
                        workshit[row_item, col_pos + 9] = po.AW;
                        workshit[row_item, col_pos + 10] = po.AX;
                        workshit[row_item, col_pos + 11] = po.AY;
                        workshit[row_item, col_pos + 12] = po.AZ;
                        workshit[row_item, col_pos + 13] = po.BA;
                        workshit[row_item, col_pos + 14] = po.BB;
                        workshit[row_item, col_pos + 15] = po.BC;

                        row_item++;
                    }

                    col_pos += 16;
                }

                gridMain.AddWorksheet(workshit);
                gridMain.CurrentWorksheet = workshit;
            }
        }

        private void AutoAppend(Worksheet sheet, int row, int col)
        {
            if (sheet.RowCount < row)
                sheet.AppendRows(row - sheet.RowCount + 1);

            if (sheet.ColumnCount < col)
                sheet.AppendCols(col - sheet.ColumnCount + 1);
        }
    }
}
