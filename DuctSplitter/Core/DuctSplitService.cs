using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace TD.DuctSplitter.Core
{
    public class SplitReport
    {
        public int DuctsProcessed;
        public int DuctsSplit;
        public int DuctsSkippedShort;
        public int BreaksCreated;
        public int UnionsCreated;
        public int JointsWithoutUnion;
        public readonly List<string> Messages = new List<string>();
        /// <summary>Mỗi phần tử là chuỗi các đoạn ống (đã sắp xếp theo thứ tự dọc tuyến) sinh ra từ 1 ống gốc.</summary>
        public readonly List<List<ElementId>> Runs = new List<List<ElementId>>();
    }

    /// <summary>
    /// Chia đốt ống gió: MechanicalUtils.BreakCurve tại các điểm cách nhau đúng chiều dài đốt,
    /// sau đó nối 2 đầu bằng Union theo Routing Preferences của kiểu ống (giống lệnh Split Duct trên giao diện).
    /// </summary>
    public class DuctSplitService
    {
        private readonly Document _doc;
        private readonly SplitSettings _s;
        private readonly HashSet<string> _noUnionTypes = new HashSet<string>();

        public DuctSplitService(Document doc, SplitSettings settings)
        {
            _doc = doc;
            _s = settings;
        }

        public double SegmentLengthFt(Duct duct)
        {
            switch (DuctHelper.Shape(duct))
            {
                case ConnectorProfileType.Round: return SplitSettings.MmToFt(_s.RoundLengthMm);
                case ConnectorProfileType.Oval: return SplitSettings.MmToFt(_s.OvalLengthMm);
                default: return SplitSettings.MmToFt(_s.RectLengthMm);
            }
        }

        /// <summary>Chạy trong một Transaction đang mở. Mỗi ống dùng SubTransaction riêng: lỗi ống nào chỉ hoàn tác ống đó.</summary>
        public SplitReport Run(IList<Duct> ducts)
        {
            var report = new SplitReport();
            foreach (Duct duct in ducts)
            {
                report.DuctsProcessed++;
                string tag = $"Ống Id {duct.Id}";

                if (duct.Pinned)
                {
                    report.Messages.Add($"{tag}: đang bị Pin – bỏ qua.");
                    continue;
                }

                using (var st = new SubTransaction(_doc))
                {
                    st.Start();
                    try
                    {
                        List<ElementId> run = SplitOne(duct, report, out bool split);
                        if (split)
                        {
                            st.Commit();
                            report.DuctsSplit++;
                            report.Runs.Add(run);
                        }
                        else
                        {
                            st.RollBack();
                            report.DuctsSkippedShort++;
                        }
                    }
                    catch (Exception ex)
                    {
                        st.RollBack();
                        report.Messages.Add($"{tag}: lỗi khi chia – {ex.Message}");
                    }
                }
            }

            foreach (string t in _noUnionTypes)
                report.Messages.Add($"Kiểu ống \"{t}\" không tạo được Union (thường do chưa khai báo Union trong Routing Preferences) → các mối chia được nối trực tiếp, dim sẽ bám vào đầu ống.");
            return report;
        }

        private List<ElementId> SplitOne(Duct duct, SplitReport report, out bool split)
        {
            split = false;
            Line line = DuctHelper.GetLine(duct);
            if (line == null) throw new InvalidOperationException("ống không phải đường thẳng.");

            double L = line.Length;
            double seg = SegmentLengthFt(duct);
            double minTail = Math.Max(SplitSettings.MmToFt(_s.MinTailMm), 2 * DuctHelper.PointTol);
            if (seg < 0.1) throw new InvalidOperationException("chiều dài đốt quá nhỏ (< 30 mm).");

            XYZ p0 = line.GetEndPoint(0), p1 = line.GetEndPoint(1);
            XYZ start = p0, end = p1;
            switch (_s.StartSide)
            {
                case StartSide.CurveEnd:
                    start = p1; end = p0; break;
                case StartSide.Auto:
                    bool c0 = DuctHelper.IsEndConnected(duct, p0);
                    bool c1 = DuctHelper.IsEndConnected(duct, p1);
                    if (c1 && !c0) { start = p1; end = p0; }
                    break;
            }
            XYZ dir = (end - start).Normalize();

            // Điểm chia: k*seg tính từ đầu bắt đầu; đốt cuối < MinTail thì gộp vào đốt trước.
            var points = new List<XYZ>();
            for (int k = 1; k * seg < L - minTail + 1e-9; k++)
                points.Add(start + dir * (k * seg));
            if (points.Count == 0) return new List<ElementId> { duct.Id };

            ElementId typeId = duct.GetTypeId();
            var pieces = new List<ElementId> { duct.Id };

            foreach (XYZ pt in points)
            {
                ElementId host = FindPieceContaining(pieces, pt);
                if (host == null) throw new InvalidOperationException("không tìm thấy đoạn ống chứa điểm chia.");
                ElementId newId = MechanicalUtils.BreakCurve(_doc, host, pt);
                pieces.Add(newId);
                _doc.Regenerate();
                report.BreaksCreated++;
            }

            // Nối các mối chia bằng Union
            foreach (XYZ pt in points)
            {
                var ends = new List<Connector>();
                foreach (ElementId id in pieces)
                {
                    if (!(_doc.GetElement(id) is Duct d)) continue;
                    Connector c = DuctHelper.ConnectorAt(d, pt);
                    if (c != null) ends.Add(c);
                }
                // Nếu BreakCurve để 2 đầu đã nối thẳng với nhau thì tách ra để chèn Union
                if (ends.Count == 2 && ends[0].IsConnectedTo(ends[1]))
                {
                    try { ends[0].DisconnectFrom(ends[1]); } catch { /* giữ nguyên */ }
                }
                ends.RemoveAll(c => c.IsConnected);
                if (ends.Count != 2)
                {
                    report.Messages.Add($"Ống Id {duct.Id}: mối chia tại {SplitSettings.FtToMm(pt.DistanceTo(start)):0} mm không xác định được 2 đầu nối.");
                    continue;
                }

                if (!_noUnionTypes.Contains(TypeName(typeId)))
                {
                    try
                    {
                        FamilyInstance u = _doc.Create.NewUnionFitting(ends[0], ends[1]);
                        if (u != null) { report.UnionsCreated++; continue; }
                    }
                    catch
                    {
                        _noUnionTypes.Add(TypeName(typeId));
                    }
                }

                // Không có Union trong Routing Preferences → nối trực tiếp để hệ thống không bị đứt
                try { ends[0].ConnectTo(ends[1]); } catch { /* để hở */ }
                report.JointsWithoutUnion++;
            }

            _doc.Regenerate();
            split = true;

            // Sắp xếp các đoạn theo thứ tự dọc tuyến từ đầu bắt đầu
            return pieces
                .Select(id => _doc.GetElement(id) as Duct)
                .Where(d => d != null && DuctHelper.GetLine(d) != null)
                .OrderBy(d => (DuctHelper.GetLine(d).Evaluate(0.5, true) - start).DotProduct(dir))
                .Select(d => d.Id)
                .ToList();
        }

        private ElementId FindPieceContaining(List<ElementId> pieces, XYZ pt)
        {
            foreach (ElementId id in pieces)
            {
                if (!(_doc.GetElement(id) is Duct d)) continue;
                Line ln = DuctHelper.GetLine(d);
                if (ln == null) continue;
                XYZ a = ln.GetEndPoint(0);
                double t = (pt - a).DotProduct(ln.Direction);
                if (t <= DuctHelper.PointTol || t >= ln.Length - DuctHelper.PointTol) continue;
                if (ln.Distance(pt) > DuctHelper.PointTol) continue;
                return id;
            }
            return null;
        }

        private string TypeName(ElementId typeId)
        {
            Element t = _doc.GetElement(typeId);
            return t?.Name ?? typeId.ToString();
        }
    }
}
