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
            // Create a new form
            Form resultForm = new Form();
            resultForm.Text = "聚合線分段分析結果";
            resultForm.Size = new Size(800, 500); // Larger window size
            resultForm.StartPosition = FormStartPosition.CenterScreen;
            resultForm.MinimizeBox = false;
            resultForm.MaximizeBox = true;
            resultForm.ShowInTaskbar = false;

            // Create a DataGridView
            DataGridView dgv = new DataGridView();
            dgv.Dock = DockStyle.Fill;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.ReadOnly = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.RowHeadersVisible = false;
            dgv.Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Regular);
            dgv.DefaultCellStyle.Font = new System.Drawing.Font("Arial", 12);
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold);

            // Calculate maximum number of segments across all data
            int maxSegments = allData.Max(data => data.AnalysisResult.Segments.Count);

            // Create a DataTable with dynamic columns
            System.Data.DataTable table = new System.Data.DataTable();
            
            // Add the first column for annotation text
            table.Columns.Add("標註文字", typeof(string));
            
            // Add columns for each segment (B1~N1 for segment numbers, B2~N2 for lengths)
            for (int i = 1; i <= maxSegments; i++)
            {
                table.Columns.Add($"線段{i}", typeof(string)); // Segment length columns
            }

            // Data rows: A2~An=標註引線中的文字, B2~N2=線段長度
            foreach (var data in allData)
            {
                var dataRow = table.NewRow();
                dataRow[0] = data.LeaderText; // A column contains annotation text
                
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

            // Bind the DataTable to the DataGridView
            dgv.DataSource = table;

            // Create an export button
            Button exportButton = new Button();
            exportButton.Text = "匯出為 CSV";
            exportButton.Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Regular);
            exportButton.Height = exportButton.Font.Height + 10; // Ensure button height is greater than text height
            exportButton.Dock = DockStyle.Bottom;
            exportButton.Click += (sender, e) =>
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV Files|*.csv";
                saveFileDialog.Title = "匯出為 CSV";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportToCsv(table, saveFileDialog.FileName);
                }
            };

            // Add the DataGridView and button to the form
            resultForm.Controls.Add(dgv);
            resultForm.Controls.Add(exportButton);

            // Show the form modally
            AcadApp.ShowModalDialog(resultForm);
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
