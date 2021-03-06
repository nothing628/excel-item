﻿using System;
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
        private List<Item> TmpItems { get; set; } = new List<Item>();

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
                    Invoke(new UpdateStatus(UpdateStat), "Importing data...Please wait", 0);
                    textBoxFileName.Text = openFileDialog.FileName;
                    btnProcess.Enabled = true;
                    btnSave.Enabled = false;

                    loadWorkSheet(openFileDialog.FileName);
                    Invoke(new UpdateStatus(UpdateStat), "Import data success!", 1);
                }
            }
        }

        private void storeToDB() {
            var startRow = Properties.Settings.Default.RowStart;
            var currentSheet = gridMain.CurrentWorksheet;
            var maxRow = currentSheet.RowCount;
            var maxCol = currentSheet.ColumnCount;
            var counter = 0;
            
            TmpItems.Clear();

            for (int i = startRow; i <= maxRow; i++)
            {
                var firstCell = currentSheet[$"L{i}"];

                if (firstCell == null || firstCell.ToString().Length == 0) //This empty data, ignore it
                    break;

                counter++;
                Invoke(new UpdateStatus(UpdateStat), $"Received Item : {counter}", 0);

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

                TmpItems.Add(item);
            }

            TmpItems = TmpItems.OrderBy(x => Convert.ToInt32(x.A)).ThenBy(x => x.AF).ToList();

            Invoke(new ClearTable(ResetGrid));

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

                gridMain.Worksheets[0].DeleteColumns(40, 1);
                gridMain.CurrentWorksheet.SetRowsHeight(3, 1, 20);
            } catch (Exception ex)
            {
                // Failed to load
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            GC.Collect();
        }

        private void buttonProcess_Click(object sender, EventArgs e)
        {
            btnProcess.Enabled = false;
            buttonChooseFile.Enabled = false;

            Invoke(new UpdateStatus(UpdateStat), "", 0);

            Task.Run(() =>
            {
                storeToDB();
                ProcessDB();
                Invoke(new Finish(FinishDB));
                Invoke(new UpdateStatus(UpdateStat), "PO Processed Complete", 1);
            });
        }

        private void FinishDB()
        {
            btnProcess.Enabled = true;
            buttonChooseFile.Enabled = true;
            btnSave.Enabled = true;
        }

        private void ProcessDB()
        {
            using (var workbook = ReoGridControl.CreateMemoryWorkbook())
            {
                try
                {
                    var workshit = workbook.CreateWorksheet("Final");
                    var items = TmpItems;
                    var grp_material = (from a in items
                                        group a by new
                                        {
                                            a.AF,
                                            a.AH,
                                            a.AL,
                                        } into b
                                        select new MaterialClass()
                                        {
                                            MatlNo = b.Key.AF,
                                            Unit = b.Key.AH,
                                            SupplierCod = b.Key.AL,
                                            ListItems = b.ToList(),
                                        }).ToList();

                    var grp_item = (from a in items
                                    group a by new { a.A, a.AF } into b
                                    select new POClass
                                    {
                                        Material = b.Key.AF,
                                        Week = b.Key.A,
                                        POs = b.ToList()
                                    }).ToList();

                    Invoke(new UpdateStatus(UpdateStat), $"Creating Header ({grp_material.Count})...", 0);
                    GenerateHeader(workshit, grp_material);
                    GenerateOutput(workshit, grp_item, grp_material);

                    workbook.AddWorksheet(workshit);
                    workbook.RemoveWorksheet(0);

                    Invoke(new UpdateTables(UpdateWorkbook), workbook, true);
                }
                catch (Exception ex)
                {
                    Invoke(new UpdateTables(UpdateWorkbook), null, false);
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            GC.Collect();
        }

        private int GetColumnPosition(List<MaterialClass> list_matl, string MatNo, string SupplierCod, string Unit)
        {
            var material = (from a in list_matl
                            where a.MatlNo == MatNo
                            where a.SupplierCod == SupplierCod
                            where a.Unit == Unit
                            select a).FirstOrDefault();

            if (material != null)
                return list_matl.IndexOf(material) * 16 + 23;
            return -1;
        }

        private void GenerateHeader(Worksheet sheet, List<MaterialClass> list_matl)
        {
            var row_pos = 1;
            var col_pos = 23; // Start at Column 'X'
            var side_style = new WorksheetRangeStyle()
            {
                Flag = PlainStyleFlag.BackColor,
                BackColor = Color.Aqua,
            };

            var align = new WorksheetRangeStyle()
            {
                Flag = PlainStyleFlag.HorizontalAlign,
                HAlign = ReoGridHorAlign.Left,
            };

            var header_style = new WorksheetRangeStyle()
            {
                Flag = PlainStyleFlag.BackColor,
                BackColor = Color.Yellow,
            };

            sheet[row_pos + 12, 0] = "週次 week";
            sheet[row_pos + 12, 1] = "投產周別 production week";
            sheet[row_pos + 12, 2] = "投產日期 production date";
            sheet[row_pos + 12, 3] = "出貨週別 shipment week";
            sheet[row_pos + 12, 4] = "工廠預計出貨日 Factory shipment date";
            sheet[row_pos + 12, 5] = "每周計劃工作時數 weekly working hours";
            sheet[row_pos + 12, 6] = "每週全廠現有生產線數(中線140)";
            sheet[row_pos + 12, 7] = "每週全廠現有每小時產能";
            sheet[row_pos + 12, 8] = "每週全廠產能(效率100%)";
            sheet[row_pos + 12, 9] = "有效每週產能 (95%)";
            sheet[row_pos + 12, 10] = "累計週次有效產能 (95%)";
            sheet[row_pos + 12, 11] = "PO#";
            sheet[row_pos + 12, 12] = "Model No";
            sheet[row_pos + 12, 13] = "楦頭(Last No)";
            sheet[row_pos + 12, 14] = "Model Name";
            sheet[row_pos + 12, 15] = "底模(Tooling No)";
            sheet[row_pos + 12, 16] = "刀模(Upper ID)";
            sheet[row_pos + 12, 17] = "MI#";
            sheet[row_pos + 12, 18] = "MI Group MI 組別";
            sheet[row_pos + 12, 19] = "顏色(Article)";
            sheet[row_pos + 12, 20] = "Order Qty";
            sheet[row_pos + 12, 21] = "累計訂單數量";
            sheet[row_pos + 12, 22] = "料號";

            sheet.SetRangeBorders(row_pos + 12, 0, 1, 23, BorderPositions.All, new RangeBorderStyle(Color.Black, BorderLineStyle.Solid));

            var col_count = list_matl.Count * 16 + 23;

            if (col_count > 16384)
                throw new Exception("This excel has reached the column limit, Please decrease the count of data! (This will have " + col_count + " Columns)");

            foreach (var item in list_matl)
            {
                // Group print
                AutoAppend(sheet, row_pos + 12, col_pos + 15);

                sheet[row_pos + 0, col_pos] = "料號";
                sheet[row_pos + 1, col_pos] = "料名";
                sheet[row_pos + 2, col_pos] = "規格";
                sheet[row_pos + 3, col_pos] = "單位";
                sheet[row_pos + 4, col_pos] = "供應商";
                sheet[row_pos + 5, col_pos] = "MOQ";
                sheet[row_pos + 6, col_pos] = "交期天數";
                sheet[row_pos + 7, col_pos] = "運輸天數";
                sheet[row_pos + 8, col_pos] = "提前加工天數";
                sheet[row_pos + 9, col_pos] = "安心庫存天數";
                sheet[row_pos + 10, col_pos] = "下單天數";

                sheet[row_pos + 0, col_pos + 1] = item.MatlNo;
                sheet[row_pos + 1, col_pos + 1] = item.MatlNa;
                sheet[row_pos + 3, col_pos + 1] = item.Unit;
                sheet[row_pos + 4, col_pos + 1] = item.SupplierCod + " " + item.SupplierNa;
                sheet[row_pos + 6, col_pos + 1] = item.MatlLT;
                sheet[row_pos + 7, col_pos + 1] = item.TransDat;

                sheet.MergeRange(row_pos + 0, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 1, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 2, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 3, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 4, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 5, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 6, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 7, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 8, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 9, col_pos + 1, 1, 15);
                sheet.MergeRange(row_pos + 10, col_pos + 1, 1, 15);

                sheet[row_pos + 12, col_pos + 0] = "單位用量";
                sheet[row_pos + 12, col_pos + 1] = "損耗率";
                sheet[row_pos + 12, col_pos + 2] = "需求量";
                sheet[row_pos + 12, col_pos + 3] = "已採購數量";
                sheet[row_pos + 12, col_pos + 4] = "採購單號";
                sheet[row_pos + 12, col_pos + 5] = "採購日期";
                sheet[row_pos + 12, col_pos + 6] = "已發出數量";
                sheet[row_pos + 12, col_pos + 7] = "領料單號";
                sheet[row_pos + 12, col_pos + 8] = "領料日期";
                sheet[row_pos + 12, col_pos + 9] = "採購數量(未交量)";
                sheet[row_pos + 12, col_pos + 10] = "訂單號碼(採購單號)";
                sheet[row_pos + 12, col_pos + 11] = "入庫日期";
                sheet[row_pos + 12, col_pos + 12] = "數量";
                sheet[row_pos + 12, col_pos + 13] = "單號";
                sheet[row_pos + 12, col_pos + 14] = "出庫日期";
                sheet[row_pos + 12, col_pos + 15] = "期末庫存";

                sheet.SetRangeBorders(row_pos, col_pos, 11, 16, BorderPositions.All, new RangeBorderStyle(Color.Black, BorderLineStyle.Solid));
                sheet.SetRangeBorders(row_pos + 12, col_pos, 1, 16, BorderPositions.All, new RangeBorderStyle(Color.Black, BorderLineStyle.Solid));
                sheet.SetRangeStyles(row_pos, col_pos, 2, 1, side_style);
                sheet.SetRangeStyles(row_pos + 3, col_pos, 2, 1, side_style);
                sheet.SetRangeStyles(row_pos + 6, col_pos, 2, 1, side_style);
                sheet.SetRangeStyles(row_pos, col_pos, 11, 15, align);
                sheet.SetRangeStyles(row_pos + 8, col_pos + 1, 2, 15, header_style);
                col_pos += 16;
            }
        }

        private void GenerateOutput(Worksheet sheet, List<POClass> list_po, List<MaterialClass> list_matl)
        {
            var item_pos = 1;
            var row_item = 14;

            foreach (var item in list_po)
            {
                AutoAppend(sheet, row_item + item.POs.Count - 1, -1);

                foreach (var detPo in item.POs)
                {
                    var col_x = GetColumnPosition(list_matl, detPo.AF, detPo.AL, detPo.AH);// GetColumn position by material and supplier;

                    if (col_x == -1) throw new Exception("Material Not Found!");

                    sheet[row_item, 0] = detPo.A;
                    sheet[row_item, 1] = "'" + detPo.B;
                    sheet[row_item, 2] = detPo.C;
                    sheet[row_item, 3] = "'" + detPo.D;
                    sheet[row_item, 4] = detPo.E;
                    sheet[row_item, 5] = detPo.F;
                    sheet[row_item, 6] = detPo.G;
                    sheet[row_item, 7] = detPo.H;
                    sheet[row_item, 8] = detPo.I;
                    sheet[row_item, 9] = detPo.J;
                    sheet[row_item, 10] = detPo.K;
                    sheet[row_item, 11] = "'" + detPo.L;
                    sheet[row_item, 12] = detPo.M;
                    sheet[row_item, 13] = "'" + detPo.N;
                    sheet[row_item, 14] = detPo.O;
                    sheet[row_item, 15] = "'" + detPo.P;
                    sheet[row_item, 16] = "'" + detPo.Q;
                    sheet[row_item, 17] = "'" + detPo.R;
                    sheet[row_item, 18] = detPo.S;
                    sheet[row_item, 19] = detPo.T;
                    sheet[row_item, 20] = detPo.U;
                    sheet[row_item, 21] = detPo.V;
                    sheet[row_item, 22] = detPo.AF;
                    sheet[row_item, col_x + 0] = detPo.AI;
                    sheet[row_item, col_x + 1] = detPo.AJ;
                    sheet[row_item, col_x + 2] = detPo.AK;
                    sheet[row_item, col_x + 3] = detPo.AQ;
                    sheet[row_item, col_x + 4] = detPo.AR;
                    sheet[row_item, col_x + 5] = detPo.AS;
                    sheet[row_item, col_x + 6] = detPo.AT;
                    sheet[row_item, col_x + 7] = detPo.AU;
                    sheet[row_item, col_x + 8] = detPo.AV;
                    sheet[row_item, col_x + 9] = detPo.AW;
                    sheet[row_item, col_x + 10] = detPo.AX;
                    sheet[row_item, col_x + 11] = detPo.AY;
                    sheet[row_item, col_x + 12] = detPo.AZ;
                    sheet[row_item, col_x + 13] = detPo.BA;
                    sheet[row_item, col_x + 14] = detPo.BB;
                    sheet[row_item, col_x + 15] = detPo.BC;

                    sheet.SetRangeBorders(row_item, col_x, 1, 16, BorderPositions.All, new RangeBorderStyle(Color.Black, BorderLineStyle.Solid));

                    row_item++;
                }
                
                var progress = Convert.ToSingle(item_pos) / Convert.ToSingle(list_po.Count);

                Invoke(new UpdateStatus(UpdateStat), $"PO Processed {item_pos} / { list_po.Count }", progress);
                
                item_pos++;
            }

            sheet.SetRangeBorders(14, 0, row_item - 14, 23, BorderPositions.All, new RangeBorderStyle(Color.Black, BorderLineStyle.Solid));
        }

        private delegate void UpdateStatus(string status, float progress);
        private delegate void UpdateTables(IWorkbook workbook, bool IsSuccess);
        private delegate void ClearTable();
        private delegate void Finish();

        private void UpdateWorkbook(IWorkbook workbook, bool IsSuccess)
        {
            if (IsSuccess)
            {
                var wrkSht = workbook.Worksheets[0];

                workbook.RemoveWorksheet(wrkSht);
                
                gridMain.AddWorksheet(wrkSht);
                gridMain.CurrentWorksheet = wrkSht;
                gridMain.RemoveWorksheet(0);
            }
        }

        private void UpdateStat(string status, float progress)
        {
            toolStripProgressBar1.Value = Convert.ToInt32(progress * 100);
            toolStripStatusLabel1.Text = "Status : " + status;
        }

        private void ResetGrid()
        {
            gridMain.Reset();
            GC.Collect();
        }

        private void AutoAppend(Worksheet sheet, int row, int col)
        {
            if (sheet.RowCount < row + 1)
                sheet.AppendRows(row - sheet.RowCount + 1);

            if (sheet.ColumnCount < col + 1)
                sheet.AppendCols(col - sheet.ColumnCount + 1);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            using (var saveDlg = new SaveFileDialog() { Filter = "Excel File|*.xlsx"})
            {
                if (saveDlg.ShowDialog() == DialogResult.OK)
                {
                    Invoke(new UpdateStatus(UpdateStat), "Exporting data...Please wait", 0);
                    gridMain.Save(saveDlg.FileName, unvell.ReoGrid.IO.FileFormat.Excel2007, Encoding.Unicode);
                    btnSave.Enabled = false;
                    btnProcess.Enabled = false;
                    Invoke(new UpdateStatus(UpdateStat), "Export data success!", 1);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var sett = new Setting())
            {
                if (sett.ShowDialog() == DialogResult.OK)
                {
                    //
                }
            }
        }
    }

    public class MaterialClass
    {
        public string MatlNo { get; set; }
        public string MatlNa
        {
            get
            {
                return ListItems.First().AG;
            }
        }
        public string Unit { get; set; }
        public string SupplierCod { get; set; }
        public string SupplierNa
        {
            get {
                return ListItems.First().AM;
            }
        }
        public string MatlLT
        {
            get
            {
                var lst = new List<int>();
                var matLt = (from a in ListItems
                             select a.AO).Distinct().ToList();
                var str = "";

                matLt.ForEach(x => {
                    try
                    {
                        var s = Convert.ToInt32(x);

                        if (s > 0)
                            lst.Add(s);
                    } catch { }
                });

                lst.ForEach(x => str += x.ToString() + ", ");

                return str.Substring(0, str.Length - 2);
            }
        }
        public string TransDat
        {
            get
            {
                var lst = new List<int>();
                var matLt = (from a in ListItems
                             select a.AP).Distinct().ToList();
                var str = "";

                matLt.ForEach(x => {
                    try
                    {
                        var s = Convert.ToInt32(x);

                        if (s > 0)
                            lst.Add(s);
                    }
                    catch { }
                });

                lst.ForEach(x => str += x.ToString() + ", ");

                return str.Substring(0, str.Length - 2);
            }
        }
        public List<Item> ListItems { get; set; }
    }

    public class POClass
    {
        public string Week { get; set; }
        public string Material { get; set; }
        public List<Item> POs { get; set; }
    }
}
