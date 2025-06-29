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
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCAD_SumDim.MyCommands))]
[assembly: ExtensionApplication(typeof(AutoCAD_SumDim.PluginExtension))]

namespace AutoCAD_SumDim
{
    public class MyCommands
    {
        // 聚合線分段分析命令 - 支援框選和點選
        [CommandMethod("PLPICK", "聚合線點選", "聚合線點選", CommandFlags.Modal)]
        public void PolylinePickStats()
        {
            try
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                ed.WriteMessage("\n請選擇聚合線進行分析:");
                ed.WriteMessage("\n提示: 可以框選多條聚合線，或逐一點選聚合線");

                // 提示用戶選擇聚合線，支援框選和點選
                List<ObjectId> selectedPolylines = new List<ObjectId>();

                // 創建選擇選項，允許框選
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\n請選擇聚合線（支援框選或點選）: ";
                pso.MessageForRemoval = "\n移除聚合線: ";
                pso.AllowDuplicates = false;

                // 創建選擇過濾器，只允許選擇聚合線
                TypedValue[] filterArray = new TypedValue[1];
                filterArray[0] = new TypedValue((int)DxfCode.Start, "LWPOLYLINE");
                SelectionFilter filter = new SelectionFilter(filterArray);

                // 執行選擇
                PromptSelectionResult psr = ed.GetSelection(pso, filter);
                
                if (psr.Status == PromptStatus.OK)
                {
                    selectedPolylines.AddRange(psr.Value.GetObjectIds());
                    ed.WriteMessage($"\n已選擇 {selectedPolylines.Count} 條聚合線");
                }
                else
                {
                    ed.WriteMessage("\n未選擇任何聚合線，操作已取消。");
                    return;
                }

                if (selectedPolylines.Count == 0)
                {
                    ed.WriteMessage("\n未選擇任何聚合線，操作已取消。");
                    return;
                }

                // 提示用戶選擇引線標註（可選）
                ed.WriteMessage("\n請選擇相關的引線標註（可選，按Enter跳過）:");
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

                // 分析聚合線
                var analyzer = new SimplePolylineAnalyzer();
                var result = analyzer.AnalyzePolylines(selectedPolylines, leaderTexts);

                if (result != null)
                {
                    // 顯示結果彈窗
                    ShowResultDialog(result);
                }
                else
                {
                    ed.WriteMessage("\n無法分析選擇的聚合線。");
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

        private void ShowResultDialog(MultiPolylineAnalysisResult result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== 聚合線分段分析結果 ===");
            sb.AppendLine();
            
            if (result.LeaderTexts.Count > 0)
            {
                sb.AppendLine("引線標註文字:");
                foreach (var text in result.LeaderTexts)
                {
                    sb.AppendLine($"  • {text}");
                }
                sb.AppendLine();
            }
            
            sb.AppendLine("線段分析:");
            for (int i = 0; i < result.Segments.Count; i++)
            {
                var segment = result.Segments[i];
                sb.AppendLine($"  第 {i + 1} 段: {segment.Length:F2} 單位");
            }
            sb.AppendLine();
            
            sb.AppendLine($"總長度: {result.TotalLength:F2} 單位");
            sb.AppendLine($"總段數: {result.Segments.Count} 段");
            sb.AppendLine($"聚合線數量: {result.PolylineCount} 條");

            MessageBox.Show(sb.ToString(), "聚合線分析結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
