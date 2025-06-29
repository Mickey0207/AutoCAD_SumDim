using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

namespace AutoCAD_SumDim
{
    // 多條聚合線分析結果類別
    public class MultiPolylineAnalysisResult
    {
        public List<PolylineSegmentInfo> Segments { get; set; } = new List<PolylineSegmentInfo>();
        public List<string> LeaderTexts { get; set; } = new List<string>();
        public double TotalLength => Segments.Sum(s => s.Length);
        public int PolylineCount { get; set; }
    }

    // 單條聚合線分析結果類別
    public class PolylineAnalysisResult
    {
        public List<PolylineSegmentInfo> Segments { get; set; } = new List<PolylineSegmentInfo>();
        public List<string> LeaderTexts { get; set; } = new List<string>();
        public double TotalLength => Segments.Sum(s => s.Length);
    }

    public class PolylineSegmentInfo
    {
        public Point3d StartPoint { get; set; }
        public Point3d EndPoint { get; set; }
        public double Length { get; set; }
        public string SegmentType { get; set; } = "直線"; // "直線" 或 "弧線"
    }

    // 簡化的聚合線分析器
    public class SimplePolylineAnalyzer
    {
        public MultiPolylineAnalysisResult AnalyzePolylines(List<ObjectId> polylineIds, List<string> leaderTexts)
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var result = new MultiPolylineAnalysisResult();
                    result.LeaderTexts = leaderTexts ?? new List<string>();
                    result.PolylineCount = polylineIds.Count;

                    // 分析每條聚合線的線段
                    foreach (var polylineId in polylineIds)
                    {
                        Polyline pline = tr.GetObject(polylineId, OpenMode.ForRead) as Polyline;
                        if (pline == null) continue;

                        // 分析聚合線的每個線段（每個彎折點之間為一段）
                        for (int i = 0; i < pline.NumberOfVertices - 1; i++)
                        {
                            Point3d segStartPoint = pline.GetPoint3dAt(i);
                            Point3d segEndPoint = pline.GetPoint3dAt(i + 1);
                            
                            double segLength;
                            string segType = "直線";
                            
                            // 如果是弧段，計算弧長
                            if (pline.GetSegmentType(i) == SegmentType.Arc)
                            {
                                CircularArc3d arc = pline.GetArcSegmentAt(i);
                                segLength = arc.GetLength(arc.GetParameterOf(segStartPoint), arc.GetParameterOf(segEndPoint), 1e-10);
                                segType = "弧線";
                            }
                            else
                            {
                                // 直線段
                                segLength = segStartPoint.DistanceTo(segEndPoint);
                            }

                            result.Segments.Add(new PolylineSegmentInfo
                            {
                                StartPoint = segStartPoint,
                                EndPoint = segEndPoint,
                                Length = segLength,
                                SegmentType = segType
                            });
                        }
                    }

                    tr.Commit();
                    return result;
                }
            }
            catch (System.Exception ex)
            {
                // 記錄錯誤到AutoCAD命令列
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\n分析聚合線時發生錯誤: {ex.Message}");
                return null;
            }
        }

        public PolylineAnalysisResult AnalyzePolyline(ObjectId polylineId, List<string> leaderTexts)
        {
            // 為了向後兼容，保留單條聚合線分析方法
            var multiResult = AnalyzePolylines(new List<ObjectId> { polylineId }, leaderTexts);
            if (multiResult == null) return null;

            return new PolylineAnalysisResult
            {
                Segments = multiResult.Segments,
                LeaderTexts = multiResult.LeaderTexts
            };
        }
    }
}