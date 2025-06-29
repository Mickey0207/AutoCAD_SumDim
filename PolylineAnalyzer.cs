using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

namespace AutoCAD_SumDim
{
    // �h���E�X�u���R���G���O
    public class MultiPolylineAnalysisResult
    {
        public List<PolylineSegmentInfo> Segments { get; set; } = new List<PolylineSegmentInfo>();
        public List<string> LeaderTexts { get; set; } = new List<string>();
        public double TotalLength => Segments.Sum(s => s.Length);
        public int PolylineCount { get; set; }
    }

    // ����E�X�u���R���G���O
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
        public string SegmentType { get; set; } = "���u"; // "���u" �� "���u"
    }

    // ²�ƪ��E�X�u���R��
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

                    // ���R�C���E�X�u���u�q
                    foreach (var polylineId in polylineIds)
                    {
                        Polyline pline = tr.GetObject(polylineId, OpenMode.ForRead) as Polyline;
                        if (pline == null) continue;

                        // ���R�E�X�u���C�ӽu�q�]�C���s���I�������@�q�^
                        for (int i = 0; i < pline.NumberOfVertices - 1; i++)
                        {
                            Point3d segStartPoint = pline.GetPoint3dAt(i);
                            Point3d segEndPoint = pline.GetPoint3dAt(i + 1);
                            
                            double segLength;
                            string segType = "���u";
                            
                            // �p�G�O���q�A�p�⩷��
                            if (pline.GetSegmentType(i) == SegmentType.Arc)
                            {
                                CircularArc3d arc = pline.GetArcSegmentAt(i);
                                segLength = arc.GetLength(arc.GetParameterOf(segStartPoint), arc.GetParameterOf(segEndPoint), 1e-10);
                                segType = "���u";
                            }
                            else
                            {
                                // ���u�q
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
                // �O�����~��AutoCAD�R�O�C
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\n���R�E�X�u�ɵo�Ϳ��~: {ex.Message}");
                return null;
            }
        }

        public PolylineAnalysisResult AnalyzePolyline(ObjectId polylineId, List<string> leaderTexts)
        {
            // ���F�V��ݮe�A�O�d����E�X�u���R��k
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