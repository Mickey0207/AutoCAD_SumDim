using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Data;
using System.Drawing;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using OfficeOpenXml;

[assembly: CommandClass(typeof(AutoCAD_SumDim.MyCommands))]
[assembly: ExtensionApplication(typeof(AutoCAD_SumDim.PluginExtension))]

namespace AutoCAD_SumDim
{
    public class MyCommands
    {
        // A class to hold the analysis data for each polyline
        private class PolylineAnalysisData
        {
            public string LeaderText { get; set; }
            public MultiPolylineAnalysisResult AnalysisResult { get; set; }
        }

        // 聚合線分段分析命令 - 支援重複選擇和詳細分析
        [CommandMethod("PLPICK", "聚合線點選", "PLPICK", CommandFlags.Modal)]
        public void PolylinePickStats()
        {
            try
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                // List to store all analysis data
                var allAnalysisData = new List<PolylineAnalysisData>();

                while (true)
                {
                    // 1. Prompt user to select polylines (support both window selection and picking)
                    ed.WriteMessage("\n請選擇聚合線進行分析 (可框選多條或點選單條，按Enter結束): ");
                    
                    List<ObjectId> selectedPolylines = new List<ObjectId>();

                    // 創建選擇選項，允許框選和點選
                    PromptSelectionOptions pso = new PromptSelectionOptions();
                    pso.MessageForAdding = "\n請選擇聚合線（支援框選或點選，Enter結束）: ";
                    pso.MessageForRemoval = "\n移除聚合線: ";
                    pso.AllowDuplicates = false;
                    pso.SingleOnly = false; // Allow multiple selection

                    // 創建選擇過濾器，只允許選擇聚合線
                    TypedValue[] filterArray = new TypedValue[1];
                    filterArray[0] = new TypedValue((int)DxfCode.Start, "LWPOLYLINE");
                    SelectionFilter filter = new SelectionFilter(filterArray);

                    // 執行選擇
                    PromptSelectionResult psr = ed.GetSelection(pso, filter);

                    if (psr.Status == PromptStatus.Cancel)
                    {
                        ed.WriteMessage("\n操作已取消。");
                        return;
                    }

                    if (psr.Status == PromptStatus.None || psr.Status != PromptStatus.OK)
                    {
                        // User pressed Enter or no selection, break the loop
                        break;
                    }

                    selectedPolylines.AddRange(psr.Value.GetObjectIds());
                    ed.WriteMessage($"\n已選擇 {selectedPolylines.Count} 條聚合線");

                    // 2. Prompt user to select leader texts for this group
                    ed.WriteMessage("\n請選擇對應的標註文字（可選擇多個，按Enter結束）:");
                    List<string> leaderTexts = new List<string>();
                    
                    while (true)
                    {
                        PromptEntityOptions peoLeader = new PromptEntityOptions("\n請點選引線標註或文字（Enter結束）: ");
                        peoLeader.AllowNone = true;
                        peoLeader.SetRejectMessage("\n請選擇引線、文字物件或按Enter結束!");
                        peoLeader.AddAllowedClass(typeof(Leader), true);
                        peoLeader.AddAllowedClass(typeof(MLeader), true);
                        peoLeader.AddAllowedClass(typeof(DBText), true);
                        peoLeader.AddAllowedClass(typeof(MText), true);

                        PromptEntityResult perLeader = ed.GetEntity(peoLeader);
                        if (perLeader.Status == PromptStatus.None)
                        {
                            break; // 用戶按了Enter，結束選擇
                        }
                        else if (perLeader.Status == PromptStatus.OK)
                        {
                            string text = GetEntityText(perLeader.ObjectId);
                            if (!string.IsNullOrEmpty(text))
                            {
                                leaderTexts.Add(text);
                                ed.WriteMessage($"\n已添加文字: {text}");
                            }
                        }
                        else
                        {
                            break; // 其他狀況結束選擇
                        }
                    }

                    // 3. Analyze the polylines
                    var analyzer = new SimplePolylineAnalyzer();
                    var result = analyzer.AnalyzePolylines(selectedPolylines, leaderTexts);

                    if (result != null)
                    {
                        // 4. Store and display data in command line
                        string mainText = leaderTexts.Count > 0 ? leaderTexts[0] : "聚合線分析";
                        allAnalysisData.Add(new PolylineAnalysisData { LeaderText = mainText, AnalysisResult = result });
                        ed.WriteMessage($"\n已記錄: 標註='{mainText}', 總長度={result.TotalLength:F2}, 總段數={result.Segments.Count}");
                    }
                    else
                    {
                        ed.WriteMessage("\n無法分析所選聚合線。");
                    }
                }

                // 5. Show final results in a DataGridView if any data was collected
                if (allAnalysisData.Count > 0)
                {
                    ShowNewFormatResultDialog(allAnalysisData);
                }
                else
                {
                    ed.WriteMessage("\n沒有收集到任何數據。");
                }
            }
            catch (System.Exception ex)
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\n執行聚合線分析時發生錯誤: {ex.Message}");
            }
        }

        private string GetEntityText(ObjectId entityId)
        {
            try
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                    
                    if (ent is DBText dbText)
                    {
                        return dbText.TextString;
                    }
                    else if (ent is MText mText)
                    {
                        return mText.Contents;
                    }
                    else if (ent is Leader leader)
                    {
                        // 尋找引線相關的標註文字
                        return FindLeaderText(leader);
                    }
                    else if (ent is MLeader mLeader)
                    {
                        return mLeader.MText?.Contents ?? "";
                    }
                    
                    tr.Commit();
                }
            }
            catch
            {
                // 忽略錯誤
            }
            
            return "";
        }

        private string FindLeaderText(Leader leader)
        {
            try
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 嘗試找到引線相關的標註
                    if (leader.HasArrowHead)
                    {
                        Point3d endPoint = leader.VertexAt(leader.NumVertices - 1);
                        
                        // 在引線終點附近尋找文字
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                        foreach (ObjectId objId in ms)
                        {
                            Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                            
                            if (ent is DBText dbText)
                            {
                                if (endPoint.DistanceTo(dbText.Position) <= 5.0)
                                {
                                    return dbText.TextString;
                                }
                            }
                            else if (ent is MText mText)
                            {
                                if (endPoint.DistanceTo(mText.Location) <= 5.0)
                                {
                                    return mText.Contents;
                                }
                            }
                        }
                    }
                    
                    tr.Commit();
                }
            }
            catch
            {
                // 忽略錯誤
            }
            
            return "";
        }

        private void ShowNewFormatResultDialog(List<PolylineAnalysisData> allData)
        {
            // Create a new form with increased height
            Form resultForm = new Form();
            resultForm.Text = "聚合線分段分析結果";
            resultForm.Size = new Size(1300, 800); // Increased size for better layout
            resultForm.StartPosition = FormStartPosition.CenterScreen;
            resultForm.MinimizeBox = false;
            resultForm.MaximizeBox = true;
            resultForm.ShowInTaskbar = false;

            // Create unit conversion panel with increased height to accommodate larger radio buttons
            Panel unitPanel = new Panel();
            unitPanel.Height = 160; // Increased height for larger radio buttons
            unitPanel.Dock = DockStyle.Top;
            unitPanel.BorderStyle = BorderStyle.FixedSingle;
            unitPanel.BackColor = Color.FromArgb(250, 250, 250); // Light gray background

            // Create source unit group box with better positioning
            GroupBox sourceGroupBox = new GroupBox();
            sourceGroupBox.Text = "圖面原始單位";
            sourceGroupBox.Font = new System.Drawing.Font("Microsoft JhengHei", 10, System.Drawing.FontStyle.Bold);
            sourceGroupBox.Location = new Point(20, 15);
            sourceGroupBox.Size = new Size(500, 130); // Increased height for larger radio buttons
            sourceGroupBox.ForeColor = Color.FromArgb(64, 64, 64);

            // Source unit radio buttons with increased size and moved down from title
            RadioButton srcMm = CreateStyledRadioButton("mm (毫米)", 15, 40, true);
            RadioButton srcCm = CreateStyledRadioButton("cm (公分)", 160, 40, false);
            RadioButton srcM = CreateStyledRadioButton("m (公尺)", 15, 80, false);
            RadioButton srcKm = CreateStyledRadioButton("km (公里)", 160, 80, false);

            sourceGroupBox.Controls.Add(srcMm);
            sourceGroupBox.Controls.Add(srcCm);
            sourceGroupBox.Controls.Add(srcM);
            sourceGroupBox.Controls.Add(srcKm);

            // Create target unit group box with better positioning
            GroupBox targetGroupBox = new GroupBox();
            targetGroupBox.Text = "轉換目標單位";
            targetGroupBox.Font = new System.Drawing.Font("Microsoft JhengHei", 10, System.Drawing.FontStyle.Bold);
            targetGroupBox.Location = new Point(540, 15);
            targetGroupBox.Size = new Size(500, 130); // Increased height for larger radio buttons
            targetGroupBox.ForeColor = Color.FromArgb(64, 64, 64);

            // Target unit radio buttons with increased size and moved down from title
            RadioButton tgtMm = CreateStyledRadioButton("mm (毫米)", 15, 40, false);
            RadioButton tgtCm = CreateStyledRadioButton("cm (公分)", 160, 40, false);
            RadioButton tgtM = CreateStyledRadioButton("m (公尺)", 15, 80, true);
            RadioButton tgtKm = CreateStyledRadioButton("km (公里)", 160, 80, false);

            targetGroupBox.Controls.Add(tgtMm);
            targetGroupBox.Controls.Add(tgtCm);
            targetGroupBox.Controls.Add(tgtM);
            targetGroupBox.Controls.Add(tgtKm);

            // Create styled convert button with better positioning
            Button convertButton = new Button();
            convertButton.Text = "轉換單位";
            convertButton.Font = new System.Drawing.Font("Microsoft JhengHei", 12, System.Drawing.FontStyle.Bold);
            convertButton.Size = new Size(140, 50);
            convertButton.Location = new Point(1070, 55); // Adjusted for larger panel
            convertButton.BackColor = Color.FromArgb(0, 122, 204);
            convertButton.ForeColor = Color.White;
            convertButton.FlatStyle = FlatStyle.Flat;
            convertButton.FlatAppearance.BorderSize = 0;
            convertButton.Cursor = Cursors.Hand;

            // Add hover effects
            convertButton.MouseEnter += (s, e) => convertButton.BackColor = Color.FromArgb(0, 102, 184);
            convertButton.MouseLeave += (s, e) => convertButton.BackColor = Color.FromArgb(0, 122, 204);

            // Add all controls to unit panel
            unitPanel.Controls.Add(sourceGroupBox);
            unitPanel.Controls.Add(targetGroupBox);
            unitPanel.Controls.Add(convertButton);

            // Create a DataGridView with explicit row height of 70
            DataGridView dgv = new DataGridView();
            dgv.Dock = DockStyle.Fill;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.ReadOnly = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.RowHeadersVisible = false;
            dgv.Font = new System.Drawing.Font("Microsoft JhengHei", 11, System.Drawing.FontStyle.Regular);
            dgv.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft JhengHei", 11);
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Microsoft JhengHei", 11, System.Drawing.FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(64, 64, 64);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersHeight = 70; // Set header height to 70 pixels
            dgv.RowTemplate.Height = 70; // Set row height to exactly 70 pixels
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(173, 216, 230);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            dgv.GridColor = Color.FromArgb(200, 200, 200);
            dgv.BorderStyle = BorderStyle.None;
            dgv.DefaultCellStyle.Padding = new Padding(5); // Add padding for better readability

            // Calculate maximum number of segments across all data
            int maxSegments = allData.Max(data => data.AnalysisResult.Segments.Count);

            // Create a DataTable with dynamic columns
            System.Data.DataTable table = CreateDataTable(allData, maxSegments);

            // Bind the DataTable to the DataGridView
            dgv.DataSource = table;

            // Ensure all rows have height of 70 after data binding
            dgv.DataBindingComplete += (sender, e) =>
            {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    row.Height = 70; // Explicitly set each row height to 70
                }
            };

            // Convert button click event
            convertButton.Click += (sender, e) =>
            {
                string sourceUnit = GetSelectedUnit(srcMm, srcCm, srcM, srcKm);
                string targetUnit = GetSelectedUnit(tgtMm, tgtCm, tgtM, tgtKm);
                
                if (sourceUnit != null && targetUnit != null)
                {
                    var convertedTable = ConvertTableUnits(allData, maxSegments, sourceUnit, targetUnit);
                    dgv.DataSource = convertedTable;
                }
            };

            // Create styled export button with increased height
            Button exportButton = new Button();
            exportButton.Text = "匯出為 CSV";
            exportButton.Font = new System.Drawing.Font("Microsoft JhengHei", 12, System.Drawing.FontStyle.Bold);
            exportButton.Height = 60; // Increased height
            exportButton.Dock = DockStyle.Bottom;
            exportButton.BackColor = Color.FromArgb(46, 125, 50);
            exportButton.ForeColor = Color.White;
            exportButton.FlatStyle = FlatStyle.Flat;
            exportButton.FlatAppearance.BorderSize = 0;
            exportButton.Cursor = Cursors.Hand;

            // Add hover effects for export button
            exportButton.MouseEnter += (s, e) => exportButton.BackColor = Color.FromArgb(56, 142, 60);
            exportButton.MouseLeave += (s, e) => exportButton.BackColor = Color.FromArgb(46, 125, 50);

            exportButton.Click += (sender, e) =>
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV Files|*.csv";
                saveFileDialog.Title = "匯出為 CSV";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportToCsv((System.Data.DataTable)dgv.DataSource, saveFileDialog.FileName);
                }
            };

            // Add controls to the form
            resultForm.Controls.Add(dgv);
            resultForm.Controls.Add(unitPanel);
            resultForm.Controls.Add(exportButton);

            // Show the form modally
            AcadApp.ShowModalDialog(resultForm);
        }

        private RadioButton CreateStyledRadioButton(string text, int x, int y, bool isChecked)
        {
            RadioButton rb = new RadioButton();
            rb.Text = text;
            rb.Font = new System.Drawing.Font("Microsoft JhengHei", 11, System.Drawing.FontStyle.Regular); // Slightly larger font
            rb.Location = new Point(x, y);
            rb.Checked = isChecked;
            rb.Size = new Size(140, 35); // Increased size for larger checkbox and better readability
            rb.ForeColor = Color.FromArgb(64, 64, 64);
            
            return rb;
        }

        private string GetSelectedUnit(params RadioButton[] radioButtons)
        {
            foreach (var rb in radioButtons)
            {
                if (rb.Checked)
                {
                    if (rb.Text.Contains("mm")) return "mm";
                    if (rb.Text.Contains("cm")) return "cm";
                    if (rb.Text.Contains("m (公尺)")) return "m";
                    if (rb.Text.Contains("km")) return "km";
                }
            }
            return null;
        }

        private double GetConversionFactor(string fromUnit, string toUnit)
        {
            // Conversion factors to meters
            var toMeters = new Dictionary<string, double>
            {
                { "mm", 0.001 },
                { "cm", 0.01 },
                { "m", 1.0 },
                { "km", 1000.0 }
            };

            return toMeters[fromUnit] / toMeters[toUnit];
        }

        private System.Data.DataTable CreateDataTable(List<PolylineAnalysisData> allData, int maxSegments)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            
            // Add the first column for annotation text
            table.Columns.Add("標註文字", typeof(string));
            
            // Add columns for each segment
            for (int i = 1; i <= maxSegments; i++)
            {
                table.Columns.Add($"線段{i}", typeof(string));
            }

            // Data rows
            foreach (var data in allData)
            {
                var dataRow = table.NewRow();
                dataRow[0] = data.LeaderText;
                
                // Fill in the segment lengths
                for (int i = 0; i < data.AnalysisResult.Segments.Count; i++)
                {
                    if (i + 1 < table.Columns.Count)
                    {
                        dataRow[i + 1] = Math.Round(data.AnalysisResult.Segments[i].Length, 2).ToString();
                    }
                }
                
                // Fill remaining columns with empty strings
                for (int i = data.AnalysisResult.Segments.Count + 1; i < table.Columns.Count; i++)
                {
                    dataRow[i] = "";
                }
                
                table.Rows.Add(dataRow);

                // Add additional rows for any extra leader texts
                for (int j = 1; j < data.AnalysisResult.LeaderTexts.Count; j++)
                {
                    var extraTextRow = table.NewRow();
                    extraTextRow[0] = data.AnalysisResult.LeaderTexts[j];
                    for (int k = 1; k < table.Columns.Count; k++)
                    {
                        extraTextRow[k] = "";
                    }
                    table.Rows.Add(extraTextRow);
                }
            }

            return table;
        }

        private System.Data.DataTable ConvertTableUnits(List<PolylineAnalysisData> allData, int maxSegments, string sourceUnit, string targetUnit)
        {
            double conversionFactor = GetConversionFactor(sourceUnit, targetUnit);
            
            System.Data.DataTable table = new System.Data.DataTable();
            
            // Add the first column for annotation text
            table.Columns.Add("標註文字", typeof(string));
            
            // Add columns for each segment with unit suffix
            for (int i = 1; i <= maxSegments; i++)
            {
                table.Columns.Add($"線段{i}({targetUnit})", typeof(string));
            }

            // Data rows with converted values
            foreach (var data in allData)
            {
                var dataRow = table.NewRow();
                dataRow[0] = data.LeaderText;
                
                // Fill in the converted segment lengths
                for (int i = 0; i < data.AnalysisResult.Segments.Count; i++)
                {
                    if (i + 1 < table.Columns.Count)
                    {
                        double convertedLength = data.AnalysisResult.Segments[i].Length * conversionFactor;
                        dataRow[i + 1] = Math.Round(convertedLength, 2).ToString();
                    }
                }
                
                // Fill remaining columns with empty strings
                for (int i = data.AnalysisResult.Segments.Count + 1; i < table.Columns.Count; i++)
                {
                    dataRow[i] = "";
                }
                
                table.Rows.Add(dataRow);

                // Add additional rows for any extra leader texts
                for (int j = 1; j < data.AnalysisResult.LeaderTexts.Count; j++)
                {
                    var extraTextRow = table.NewRow();
                    extraTextRow[0] = data.AnalysisResult.LeaderTexts[j];
                    for (int k = 1; k < table.Columns.Count; k++)
                    {
                        extraTextRow[k] = "";
                    }
                    table.Rows.Add(extraTextRow);
                }
            }

            return table;
        }

        private void ExportToCsv(System.Data.DataTable table, string filePath)
        {
            using (var writer = new System.IO.StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                // Write column headers
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    writer.Write(table.Columns[col].ColumnName);
                    if (col < table.Columns.Count - 1)
                    {
                        writer.Write(",");
                    }
                }
                writer.WriteLine();

                // Write rows
                foreach (System.Data.DataRow row in table.Rows)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        writer.Write(row[col]?.ToString());
                        if (col < table.Columns.Count - 1)
                        {
                            writer.Write(",");
                        }
                    }
                    writer.WriteLine();
                }
            }
        }
    }
}
