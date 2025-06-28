using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

namespace AutoCAD_SumDim
{
    public class PolylineAnalyzer
    {
        public class PolylineSegment
        {
            public Point3d StartPoint { get; set; }
            public Point3d EndPoint { get; set; }
            public double Length { get; set; }
        }

        public class PolylineInfo
        {
            public ObjectId PolylineId { get; set; }
            public string StartBlockName { get; set; } = "";
            public string EndBlockName { get; set; } = "";
            public string StartBlockText { get; set; } = "";
            public string EndBlockText { get; set; } = "";
            public List<PolylineSegment> Segments { get; set; } = new List<PolylineSegment>();
            public double TotalLength => Segments.Sum(s => s.Length);
        }

        private readonly Dictionary<string, string> blockMapping;
        private readonly double tolerance = 1.0; // 圖塊連接容差

        public PolylineAnalyzer(Dictionary<string, string> blockMapping)
        {
            this.blockMapping = blockMapping ?? new Dictionary<string, string>();
        }

        public List<PolylineInfo> AnalyzePolylines(List<string> layerNames)
        {
            var results = new List<PolylineInfo>();
            
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                // 找出所有指定圖層的聚合線
                var polylines = new List<Polyline>();
                foreach (ObjectId objId in ms)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                    if (ent is Polyline pline && layerNames.Contains(ent.Layer))
                    {
                        polylines.Add(pline);
                    }
                }

                // 找出所有圖塊
                var blocks = new List<BlockReference>();
                foreach (ObjectId objId in ms)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                    if (ent is BlockReference blockRef)
                    {
                        blocks.Add(blockRef);
                    }
                }

                // 分析每條聚合線
                foreach (var pline in polylines)
                {
                    var polylineInfo = AnalyzePolyline(pline, blocks, tr);
                    if (polylineInfo != null)
                    {
                        results.Add(polylineInfo);
                    }
                }

                tr.Commit();
            }

            return results;
        }

        private PolylineInfo AnalyzePolyline(Polyline pline, List<BlockReference> blocks, Transaction tr)
        {
            var info = new PolylineInfo
            {
                PolylineId = pline.ObjectId
            };

            // 計算聚合線的起點和終點
            Point3d startPoint = pline.GetPoint3dAt(0);
            Point3d endPoint = pline.GetPoint3dAt(pline.NumberOfVertices - 1);

            // 尋找起點和終點連接的圖塊
            info.StartBlockName = FindNearestBlock(startPoint, blocks, out string startText);
            info.EndBlockName = FindNearestBlock(endPoint, blocks, out string endText);
            info.StartBlockText = startText;
            info.EndBlockText = endText;

            // 應用圖塊名稱映射
            if (blockMapping.ContainsKey(info.StartBlockName))
                info.StartBlockName = blockMapping[info.StartBlockName];
            if (blockMapping.ContainsKey(info.EndBlockName))
                info.EndBlockName = blockMapping[info.EndBlockName];

            // 分析聚合線的每個線段（每個彎折點之間為一段）
            for (int i = 0; i < pline.NumberOfVertices - 1; i++)
            {
                Point3d segStartPoint = pline.GetPoint3dAt(i);
                Point3d segEndPoint = pline.GetPoint3dAt(i + 1);
                
                double segLength;
                
                // 如果是弧段，計算弧長
                if (pline.GetSegmentType(i) == SegmentType.Arc)
                {
                    CircularArc3d arc = pline.GetArcSegmentAt(i);
                    segLength = arc.GetLength(arc.GetParameterOf(segStartPoint), arc.GetParameterOf(segEndPoint), 1e-10);
                }
                else
                {
                    // 直線段
                    segLength = segStartPoint.DistanceTo(segEndPoint);
                }

                info.Segments.Add(new PolylineSegment
                {
                    StartPoint = segStartPoint,
                    EndPoint = segEndPoint,
                    Length = segLength
                });
            }

            return info;
        }

        private string FindNearestBlock(Point3d point, List<BlockReference> blocks, out string blockText)
        {
            blockText = "";
            string nearestBlockName = "";
            double minDistance = double.MaxValue;

            foreach (var block in blocks)
            {
                double distance = point.DistanceTo(block.Position);
                if (distance < tolerance && distance < minDistance)
                {
                    minDistance = distance;
                    nearestBlockName = block.Name;
                    
                    // 嘗試找到與圖塊相關的文字或引線文字
                    blockText = FindBlockText(block);
                }
            }

            return nearestBlockName;
        }

        private string FindBlockText(BlockReference blockRef)
        {
            // 這裡可以實作尋找引線文字的邏輯
            // 暫時返回空字串，可以後續擴展
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // 在圖塊附近尋找文字物件
                    Point3d blockPos = blockRef.Position;
                    double searchRadius = 10.0; // 搜尋半徑

                    foreach (ObjectId objId in ms)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        
                        if (ent is DBText dbText)
                        {
                            if (blockPos.DistanceTo(dbText.Position) <= searchRadius)
                            {
                                return dbText.TextString;
                            }
                        }
                        else if (ent is MText mText)
                        {
                            if (blockPos.DistanceTo(mText.Location) <= searchRadius)
                            {
                                return mText.Contents;
                            }
                        }
                    }
                    
                    tr.Commit();
                }
            }
            catch
            {
                // 忽略錯誤，返回空字串
            }

            return "";
        }

        public void ExportToExcel(List<PolylineInfo> polylineInfos, string filePath)
        {
            try
            {
                // 首先嘗試使用 CSV 格式，這是最穩定的方案
                string csvPath = Path.ChangeExtension(filePath, ".csv");
                ExportToCsv(polylineInfos, csvPath);
                
                // 通知使用者已儲存為 CSV 格式
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\n資料已儲存為 CSV 格式: {csvPath}");
                ed.WriteMessage("\nCSV 檔案可以使用 Excel 開啟");
                
                // 嘗試將路徑更新為 CSV 檔案以便後續開啟
                throw new System.Exception($"資料已儲存為 CSV 格式: {csvPath}");
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }

        private void ExportToCsv(List<PolylineInfo> polylineInfos, string filePath)
        {
            var csv = new StringBuilder();
            
            // 找出最大段數
            int maxSegments = polylineInfos.Count > 0 ? polylineInfos.Max(p => p.Segments.Count) : 0;
            
            // 寫入標題行
            var headers = new List<string> { "起點圖塊", "終點圖塊", "引線文字" };
            for (int i = 1; i <= maxSegments; i++)
            {
                headers.Add($"第{i}段長度");
            }
            headers.Add("總長度");
            csv.AppendLine(string.Join(",", headers.Select(h => $"\"{h}\"")));
            
            // 寫入資料行
            foreach (var polyInfo in polylineInfos)
            {
                var row = new List<string>();
                
                row.Add($"\"{polyInfo.StartBlockName}\"");
                row.Add($"\"{polyInfo.EndBlockName}\"");
                
                // 引線文字（優先顯示起點，若無則顯示終點）
                string leaderText = !string.IsNullOrEmpty(polyInfo.StartBlockText) 
                    ? polyInfo.StartBlockText 
                    : polyInfo.EndBlockText;
                row.Add($"\"{leaderText}\"");
                
                // 各段長度
                for (int i = 0; i < maxSegments; i++)
                {
                    if (i < polyInfo.Segments.Count)
                    {
                        row.Add(Math.Round(polyInfo.Segments[i].Length, 2).ToString());
                    }
                    else
                    {
                        row.Add("");
                    }
                }
                
                // 總長度
                row.Add(Math.Round(polyInfo.TotalLength, 2).ToString());
                
                csv.AppendLine(string.Join(",", row));
            }
            
            // 儲存 CSV 檔案（使用 UTF-8 BOM 以便 Excel 正確開啟）
            File.WriteAllText(filePath, csv.ToString(), new UTF8Encoding(true));
        }
    }
}