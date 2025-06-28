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
        private readonly double tolerance = 1.0; // �϶��s���e�t

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

                // ��X�Ҧ����w�ϼh���E�X�u
                var polylines = new List<Polyline>();
                foreach (ObjectId objId in ms)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                    if (ent is Polyline pline && layerNames.Contains(ent.Layer))
                    {
                        polylines.Add(pline);
                    }
                }

                // ��X�Ҧ��϶�
                var blocks = new List<BlockReference>();
                foreach (ObjectId objId in ms)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                    if (ent is BlockReference blockRef)
                    {
                        blocks.Add(blockRef);
                    }
                }

                // ���R�C���E�X�u
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

            // �p��E�X�u���_�I�M���I
            Point3d startPoint = pline.GetPoint3dAt(0);
            Point3d endPoint = pline.GetPoint3dAt(pline.NumberOfVertices - 1);

            // �M��_�I�M���I�s�����϶�
            info.StartBlockName = FindNearestBlock(startPoint, blocks, out string startText);
            info.EndBlockName = FindNearestBlock(endPoint, blocks, out string endText);
            info.StartBlockText = startText;
            info.EndBlockText = endText;

            // ���ι϶��W�٬M�g
            if (blockMapping.ContainsKey(info.StartBlockName))
                info.StartBlockName = blockMapping[info.StartBlockName];
            if (blockMapping.ContainsKey(info.EndBlockName))
                info.EndBlockName = blockMapping[info.EndBlockName];

            // ���R�E�X�u���C�ӽu�q�]�C���s���I�������@�q�^
            for (int i = 0; i < pline.NumberOfVertices - 1; i++)
            {
                Point3d segStartPoint = pline.GetPoint3dAt(i);
                Point3d segEndPoint = pline.GetPoint3dAt(i + 1);
                
                double segLength;
                
                // �p�G�O���q�A�p�⩷��
                if (pline.GetSegmentType(i) == SegmentType.Arc)
                {
                    CircularArc3d arc = pline.GetArcSegmentAt(i);
                    segLength = arc.GetLength(arc.GetParameterOf(segStartPoint), arc.GetParameterOf(segEndPoint), 1e-10);
                }
                else
                {
                    // ���u�q
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
                    
                    // ���է��P�϶���������r�Τ޽u��r
                    blockText = FindBlockText(block);
                }
            }

            return nearestBlockName;
        }

        private string FindBlockText(BlockReference blockRef)
        {
            // �o�̥i�H��@�M��޽u��r���޿�
            // �Ȯɪ�^�Ŧr��A�i�H�����X�i
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // �b�϶�����M���r����
                    Point3d blockPos = blockRef.Position;
                    double searchRadius = 10.0; // �j�M�b�|

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
                // �������~�A��^�Ŧr��
            }

            return "";
        }

        public void ExportToExcel(List<PolylineInfo> polylineInfos, string filePath)
        {
            try
            {
                // �������ըϥ� CSV �榡�A�o�O��í�w�����
                string csvPath = Path.ChangeExtension(filePath, ".csv");
                ExportToCsv(polylineInfos, csvPath);
                
                // �q���ϥΪ̤w�x�s�� CSV �榡
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\n��Ƥw�x�s�� CSV �榡: {csvPath}");
                ed.WriteMessage("\nCSV �ɮץi�H�ϥ� Excel �}��");
                
                // ���ձN���|��s�� CSV �ɮץH�K����}��
                throw new System.Exception($"��Ƥw�x�s�� CSV �榡: {csvPath}");
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }

        private void ExportToCsv(List<PolylineInfo> polylineInfos, string filePath)
        {
            var csv = new StringBuilder();
            
            // ��X�̤j�q��
            int maxSegments = polylineInfos.Count > 0 ? polylineInfos.Max(p => p.Segments.Count) : 0;
            
            // �g�J���D��
            var headers = new List<string> { "�_�I�϶�", "���I�϶�", "�޽u��r" };
            for (int i = 1; i <= maxSegments; i++)
            {
                headers.Add($"��{i}�q����");
            }
            headers.Add("�`����");
            csv.AppendLine(string.Join(",", headers.Select(h => $"\"{h}\"")));
            
            // �g�J��Ʀ�
            foreach (var polyInfo in polylineInfos)
            {
                var row = new List<string>();
                
                row.Add($"\"{polyInfo.StartBlockName}\"");
                row.Add($"\"{polyInfo.EndBlockName}\"");
                
                // �޽u��r�]�u����ܰ_�I�A�Y�L�h��ܲ��I�^
                string leaderText = !string.IsNullOrEmpty(polyInfo.StartBlockText) 
                    ? polyInfo.StartBlockText 
                    : polyInfo.EndBlockText;
                row.Add($"\"{leaderText}\"");
                
                // �U�q����
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
                
                // �`����
                row.Add(Math.Round(polyInfo.TotalLength, 2).ToString());
                
                csv.AppendLine(string.Join(",", row));
            }
            
            // �x�s CSV �ɮס]�ϥ� UTF-8 BOM �H�K Excel ���T�}�ҡ^
            File.WriteAllText(filePath, csv.ToString(), new UTF8Encoding(true));
        }
    }
}