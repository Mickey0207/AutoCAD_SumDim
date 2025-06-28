using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCAD_SumDim.MyCommands))]
[assembly: ExtensionApplication(typeof(AutoCAD_SumDim.PluginExtension))]

namespace AutoCAD_SumDim
{
    public class MyCommands
    {
        // 聚合線長度統計命令
        [CommandMethod("PLSTATS", "聚合線統計", "聚合線統計", CommandFlags.Modal)]
        public void PolylineStats()
        {
            try
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                // 顯示用戶介面
                using (var form = new PolylineStatsForm())
                {
                    var result = AcadApp.ShowModalDialog(form);
                    
                    if (result == DialogResult.OK)
                    {
                        ed.WriteMessage("\n開始分析聚合線...");
                        
                        // 創建分析器
                        var analyzer = new PolylineAnalyzer(form.BlockMapping);
                        
                        // 分析聚合線
                        var polylineInfos = analyzer.AnalyzePolylines(form.SelectedLayers);
                        
                        if (polylineInfos.Count == 0)
                        {
                            ed.WriteMessage("\n在指定的圖層中沒有找到聚合線。");
                            MessageBox.Show("在指定的圖層中沒有找到聚合線。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        
                        // 輸出到Excel/CSV
                        string fullPath = Path.Combine(form.OutputPath, form.ExcelFileName);
                        string actualOutputPath = "";
                        
                        try
                        {
                            analyzer.ExportToExcel(polylineInfos, fullPath);
                            actualOutputPath = fullPath;
                        }
                        catch (System.Exception ex)
                        {
                            // 如果輸出失敗，檢查是否有CSV檔案產生
                            string csvPath = Path.ChangeExtension(fullPath, ".csv");
                            if (File.Exists(csvPath))
                            {
                                actualOutputPath = csvPath;
                                ed.WriteMessage($"\n注意：{ex.Message}");
                            }
                            else
                            {
                                throw ex;
                            }
                        }
                        
                        ed.WriteMessage($"\n統計完成！共分析了 {polylineInfos.Count} 條聚合線。");
                        ed.WriteMessage($"\n檔案已儲存到: {actualOutputPath}");
                        
                        string fileType = Path.GetExtension(actualOutputPath).ToUpper() == ".CSV" ? "CSV" : "Excel";
                        MessageBox.Show($"統計完成！\n\n共分析了 {polylineInfos.Count} 條聚合線。\n{fileType}檔案已儲存到:\n{actualOutputPath}\n\n注意：CSV檔案可以用Excel直接開啟", 
                                      "統計完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                        // 詢問是否開啟檔案
                        if (MessageBox.Show($"是否要開啟{fileType}檔案？", "開啟檔案", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = actualOutputPath,
                                UseShellExecute = true
                            });
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\n執行聚合線統計時發生錯誤: {ex.Message}");
                MessageBox.Show($"執行聚合線統計時發生錯誤:\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // The CommandMethod attribute can be applied to any public  member 
        // function of any public class.
        // The function should take no arguments and return nothing.
        // If the method is an intance member then the enclosing class is 
        // intantiated for each document. If the member is a static member then
        // the enclosing class is NOT intantiated.
        //
        // NOTE: CommandMethod has overloads where you can provide helpid and
        // context menu.

        // Modal Command with localized name
        [CommandMethod("MyGroup", "MyCommand", "MyCommandLocal", CommandFlags.Modal)]
        public void MyCommand() // This method can have any name
        {
            // Put your command code here
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor ed;
            if (doc != null)
            {
                ed = doc.Editor;
                ed.WriteMessage("Hello, this is your first command.");

            }
        }

        // Modal Command with pickfirst selection
        [CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void MyPickFirst() // This method can have any name
        {
            PromptSelectionResult result = AcadApp.DocumentManager.MdiActiveDocument.Editor.GetSelection();
            if (result.Status == PromptStatus.OK)
            {
                // There are selected entities
                // Put your command using pickfirst set code here
            }
            else
            {
                // There are no selected entities
                // Put your command code here
            }
        }

        // Application Session Command with localized name
        [CommandMethod("MyGroup", "MySessionCmd", "MySessionCmdLocal", CommandFlags.Modal | CommandFlags.Session)]
        public void MySessionCmd() // This method can have any name
        {
            // Put your command code here
        }

        // LispFunction is similar to CommandMethod but it creates a lisp 
        // callable function. Many return types are supported not just string
        // or integer.
        [LispFunction("MyLispFunction", "MyLispFunctionLocal")]
        public int MyLispFunction(ResultBuffer args) // This method can have any name
        {
            // Put your command code here

            // Return a value to the AutoCAD Lisp Interpreter
            return 1;
        }
    }
}
