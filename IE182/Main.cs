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
            buttonProcess.Enabled = true;
        }

        private void buttonProcess_Click(object sender, EventArgs e)
        {
            //Check file Database
            if (!File.Exists("IE182.db"))
                return;

            //Open Database
            using (var db = new LiteDatabase(@"IE182.db"))
            using (var workbook = ReoGridControl.CreateMemoryWorkbook())
            {
                var workshit = workbook.Worksheets[0];
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
                workshit.AppendRows(2000);
                workshit["A290"] = "A";

                var cell = workshit["A290"];
                foreach (var item in grp_items)
                {
                    // Group
                }
            }
        }
    }
}
