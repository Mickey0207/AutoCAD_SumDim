using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCAD_SumDim
{
    public partial class PolylineStatsForm : Form
    {
        private CheckedListBox clbLayers = null!;
        private TextBox txtExcelFileName = null!;
        private TextBox txtOutputPath = null!;
        private Button btnBrowse = null!;
        private DataGridView dgvBlockMapping = null!;
        private Button btnCalculate = null!;
        private Button btnCancel = null!;
        private Label lblLayers = null!;
        private Label lblFileName = null!;
        private Label lblOutputPath = null!;
        private Label lblBlockMapping = null!;

        public List<string> SelectedLayers { get; private set; } = new List<string>();
        public string ExcelFileName { get; private set; } = "";
        public string OutputPath { get; private set; } = "";
        public Dictionary<string, string> BlockMapping { get; private set; } = new Dictionary<string, string>();

        public PolylineStatsForm()
        {
            InitializeComponent();
            LoadLayers();
            InitializeBlockMapping();
        }

        private void InitializeComponent()
        {
            this.Text = "聚合線長度統計";
            this.Size = new Size(600, 500);
            this.StartPosition = FormStartPosition.CenterParent;

            // 圖層選擇
            lblLayers = new Label
            {
                Text = "選擇聚合線圖層:",
                Location = new Point(12, 15),
                Size = new Size(120, 23)
            };

            clbLayers = new CheckedListBox
            {
                Location = new Point(12, 40),
                Size = new Size(560, 120),
                CheckOnClick = true
            };

            // Excel檔案名稱
            lblFileName = new Label
            {
                Text = "Excel檔案名稱:",
                Location = new Point(12, 175),
                Size = new Size(100, 23)
            };

            txtExcelFileName = new TextBox
            {
                Location = new Point(120, 175),
                Size = new Size(200, 23),
                Text = "聚合線統計.xlsx"
            };

            // 輸出路徑
            lblOutputPath = new Label
            {
                Text = "輸出路徑:",
                Location = new Point(12, 210),
                Size = new Size(80, 23)
            };

            txtOutputPath = new TextBox
            {
                Location = new Point(100, 210),
                Size = new Size(380, 23),
                Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            btnBrowse = new Button
            {
                Text = "瀏覽...",
                Location = new Point(490, 210),
                Size = new Size(80, 23)
            };

            // 圖塊名稱映射表格
            lblBlockMapping = new Label
            {
                Text = "圖塊名稱映射:",
                Location = new Point(12, 245),
                Size = new Size(120, 23)
            };

            dgvBlockMapping = new DataGridView
            {
                Location = new Point(12, 270),
                Size = new Size(560, 120),
                AllowUserToDeleteRows = true,
                AllowUserToAddRows = true
            };

            dgvBlockMapping.Columns.Add("OriginalName", "原始圖塊名稱");
            dgvBlockMapping.Columns.Add("NewName", "新圖塊名稱");
            dgvBlockMapping.Columns[0].Width = 270;
            dgvBlockMapping.Columns[1].Width = 270;

            // 按鈕
            btnCalculate = new Button
            {
                Text = "開始統計",
                Location = new Point(410, 420),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK
            };

            btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(500, 420),
                Size = new Size(70, 30),
                DialogResult = DialogResult.Cancel
            };

            // 事件處理
            btnBrowse.Click += BtnBrowse_Click;
            btnCalculate.Click += BtnCalculate_Click;

            // 加入控制項
            this.Controls.AddRange(new Control[] {
                lblLayers, clbLayers,
                lblFileName, txtExcelFileName,
                lblOutputPath, txtOutputPath, btnBrowse,
                lblBlockMapping, dgvBlockMapping,
                btnCalculate, btnCancel
            });
        }

        private void LoadLayers()
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    
                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        clbLayers.Items.Add(ltr.Name);
                    }
                    
                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入圖層時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeBlockMapping()
        {
            // 加入幾個範例行
            dgvBlockMapping.Rows.Add("", "");
            dgvBlockMapping.Rows.Add("", "");
            dgvBlockMapping.Rows.Add("", "");
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "選擇Excel檔案輸出位置";
                fbd.SelectedPath = txtOutputPath.Text;

                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtOutputPath.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnCalculate_Click(object sender, EventArgs e)
        {
            // 收集選擇的圖層
            SelectedLayers.Clear();
            foreach (string layer in clbLayers.CheckedItems)
            {
                SelectedLayers.Add(layer);
            }

            if (SelectedLayers.Count == 0)
            {
                MessageBox.Show("請至少選擇一個圖層", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ExcelFileName = txtExcelFileName.Text.Trim();
            if (string.IsNullOrEmpty(ExcelFileName))
            {
                MessageBox.Show("請輸入Excel檔案名稱", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!ExcelFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                ExcelFileName += ".xlsx";
            }

            OutputPath = txtOutputPath.Text.Trim();
            if (string.IsNullOrEmpty(OutputPath) || !Directory.Exists(OutputPath))
            {
                MessageBox.Show("請選擇有效的輸出路徑", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 收集圖塊映射
            BlockMapping.Clear();
            foreach (DataGridViewRow row in dgvBlockMapping.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    string original = row.Cells[0].Value.ToString().Trim();
                    string newName = row.Cells[1].Value.ToString().Trim();
                    
                    if (!string.IsNullOrEmpty(original) && !string.IsNullOrEmpty(newName))
                    {
                        BlockMapping[original] = newName;
                    }
                }
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}