using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace TD.DuctSplitter.Core
{
    /// <summary>
    /// Tạo chuỗi dim dọc tuyến ống qua các Union (và 2 đầu ống nếu chọn) trong view hiện hành.
    /// Ưu tiên tham chiếu mặt phẳng tâm của family Union vuông góc trục ống; nếu không được thì bám mặt đầu ống.
    /// </summary>
    public class UnionDimensionService
    {
        private readonly Document _doc;
        private readonly View _view;
        private readonly SplitSettings _s;
        private readonly DimensionType _dimType;

        public UnionDimensionService(Document doc, View view, SplitSettings settings)
        {
            _doc = doc;
            _view = view;
            _s = settings;
            _dimType = FindDimensionType(doc, settings.DimensionTypeName);
        }

        /// <summary>View có cho phép tạo dim không. 3D view, sheet, schedule, drafting... đều bị loại.</summary>
        public static bool ViewSupportsDims(View v, out string reason)
        {
            reason = null;
            if (v == null) { reason = "không có view hiện hành."; return false; }
            if (v is View3D) { reason = "view hiện hành là 3D – không tạo dim (chỉ chia ống)."; return false; }
            if (v.IsTemplate) { reason = "view hiện hành là View Template."; return false; }
            switch (v.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.EngineeringPlan:
                case ViewType.AreaPlan:
                case ViewType.Section:
                case ViewType.Elevation:
                case ViewType.Detail:
                    return true;
                default:
                    reason = $"view hiện hành ({v.ViewType}) không phải mặt bằng/mặt cắt – không tạo dim.";
                    return false;
            }
        }

        public static IList<DimensionType> LinearDimensionTypes(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(DimensionType))
                .Cast<DimensionType>()
                .Where(t => t.StyleType == DimensionStyleType.Linear && !string.IsNullOrEmpty(t.Name))
                .OrderBy(t => t.Name)
                .ToList();
        }

        private static DimensionType FindDimensionType(Document doc, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return LinearDimensionTypes(doc).FirstOrDefault(t => t.Name == name);
        }

        /// <summary>Tạo 1 chuỗi dim cho các đoạn ống thẳng hàng, đã sắp thứ tự. Trả về Dimension hoặc null (kèm lý do).</summary>
        public Dimension DimensionRun(IList<Duct> ordered, out string error)
        {
            error = null;
            if (ordered == null || ordered.Count < 2) { error = "chỉ có 1 đoạn ống – không có Union để dim."; return null; }

            Line first = DuctHelper.GetLine(ordered[0]);
            Line last = DuctHelper.GetLine(ordered[ordered.Count - 1]);
            if (first == null || last == null) { error = "không đọc được trục ống."; return null; }

            XYZ dir = (last.Evaluate(0.5, true) - first.Evaluate(0.5, true)).Normalize();
            XYZ viewDir = _view.ViewDirection;
            if (Math.Abs(dir.DotProduct(viewDir)) > 1e-3)
            {
                error = "ống không song song mặt phẳng view (ống đứng/ống dốc) – bỏ qua dim.";
                return null;
            }

            XYZ startPt = MinEnd(first, dir);
            XYZ endPt = MaxEnd(last, dir);

            // Phương án 1: tham chiếu Union (fallback mặt đầu ống từng mối)
            ReferenceArray refs = BuildRefs(ordered, dir, startPt, endPt, preferUnion: true);
            Dimension dim = TryCreate(ordered[0], dir, startPt, endPt, refs, out string e1);
            if (dim != null) return dim;

            // Phương án 2: toàn bộ bám mặt đầu ống
            refs = BuildRefs(ordered, dir, startPt, endPt, preferUnion: false);
            dim = TryCreate(ordered[0], dir, startPt, endPt, refs, out string e2);
            if (dim != null) return dim;

            error = e2 ?? e1 ?? "không tạo được dim.";
            return null;
        }

        private ReferenceArray BuildRefs(IList<Duct> ordered, XYZ dir, XYZ startPt, XYZ endPt, bool preferUnion)
        {
            var refs = new ReferenceArray();

            if (_s.DimIncludeEnds)
            {
                Reference r = DuctEndFaceRef(ordered[0], startPt, dir);
                if (r != null) refs.Append(r);
            }

            for (int i = 0; i < ordered.Count - 1; i++)
            {
                Line li = DuctHelper.GetLine(ordered[i]);
                XYZ joint = MaxEnd(li, dir);
                Reference r = null;
                if (preferUnion)
                {
                    FamilyInstance u = DuctHelper.UnionBetween(ordered[i], ordered[i + 1]);
                    if (u != null) r = UnionCenterRef(u, dir);
                }
                if (r == null) r = DuctEndFaceRef(ordered[i], joint, dir);
                if (r == null) r = DuctEndFaceRef(ordered[i + 1], MinEnd(DuctHelper.GetLine(ordered[i + 1]), dir), dir);
                if (r != null) refs.Append(r);
            }

            if (_s.DimIncludeEnds)
            {
                Duct lastDuct = ordered[ordered.Count - 1];
                Reference r = DuctEndFaceRef(lastDuct, endPt, dir);
                if (r != null) refs.Append(r);
            }
            return refs;
        }

        private Dimension TryCreate(Duct sample, XYZ dir, XYZ startPt, XYZ endPt, ReferenceArray refs, out string error)
        {
            error = null;
            if (refs.Size < 2) { error = "không đủ tham chiếu để dim (cần ≥ 2)."; return null; }

            XYZ side = _view.ViewDirection.CrossProduct(dir);
            if (side.GetLength() < 1e-9) { error = "không xác định được phương đặt dim."; return null; }
            side = side.Normalize();

            double offset = DuctHelper.PlanWidth(sample) / 2.0
                            + SplitSettings.MmToFt(_s.DimOffsetMm) * Math.Max(1, _view.Scale);
            XYZ a = startPt + side * offset;
            XYZ b = endPt + side * offset;
            if (a.DistanceTo(b) < 0.01) { error = "tuyến ống quá ngắn."; return null; }
            Line dimLine = Line.CreateBound(a, b);

            using (var st = new SubTransaction(_doc))
            {
                st.Start();
                try
                {
                    Dimension d = _dimType != null
                        ? _doc.Create.NewDimension(_view, dimLine, refs, _dimType)
                        : _doc.Create.NewDimension(_view, dimLine, refs);
                    if (d == null) { st.RollBack(); error = "Revit trả về dim rỗng."; return null; }
                    _doc.Regenerate();
                    st.Commit();
                    return d;
                }
                catch (Exception ex)
                {
                    st.RollBack();
                    error = ex.Message;
                    return null;
                }
            }
        }

        /// <summary>Mặt phẳng tâm của Union vuông góc với trục ống (CenterLeftRight / CenterFrontBack / CenterElevation).</summary>
        private static Reference UnionCenterRef(FamilyInstance fi, XYZ dir)
        {
            try
            {
                Transform t = fi.GetTransform();
                FamilyInstanceReferenceType type;
                if (DuctHelper.IsParallel(t.BasisX, dir, 1e-2)) type = FamilyInstanceReferenceType.CenterLeftRight;
                else if (DuctHelper.IsParallel(t.BasisY, dir, 1e-2)) type = FamilyInstanceReferenceType.CenterFrontBack;
                else if (DuctHelper.IsParallel(t.BasisZ, dir, 1e-2)) type = FamilyInstanceReferenceType.CenterElevation;
                else return null;
                return fi.GetReferences(type).FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>Mặt phẳng đầu ống (pháp tuyến // trục ống) gần điểm pt nhất.</summary>
        private Reference DuctEndFaceRef(Duct duct, XYZ pt, XYZ dir)
        {
            var optionsList = new List<Options>
            {
                new Options { ComputeReferences = true, View = _view },
                new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine }
            };

            foreach (Options opt in optionsList)
            {
                GeometryElement geom;
                try { geom = duct.get_Geometry(opt); } catch { continue; }
                if (geom == null) continue;

                Reference best = null;
                double bestDist = 0.05; // ~15 mm
                foreach (Solid solid in Solids(geom))
                {
                    foreach (Face f in solid.Faces)
                    {
                        if (!(f is PlanarFace pf) || pf.Reference == null) continue;
                        if (!DuctHelper.IsParallel(pf.FaceNormal, dir, 1e-2)) continue;
                        double dist = Math.Abs((pf.Origin - pt).DotProduct(dir));
                        if (dist < bestDist) { bestDist = dist; best = pf.Reference; }
                    }
                }
                if (best != null) return best;
            }
            return null;
        }

        private static IEnumerable<Solid> Solids(GeometryElement geom)
        {
            foreach (GeometryObject go in geom)
            {
                if (go is Solid s && s.Faces.Size > 0) yield return s;
                else if (go is GeometryInstance gi)
                {
                    foreach (Solid s2 in Solids(gi.GetInstanceGeometry())) yield return s2;
                }
            }
        }

        private static XYZ MinEnd(Line l, XYZ dir)
        {
            XYZ a = l.GetEndPoint(0), b = l.GetEndPoint(1);
            return a.DotProduct(dir) <= b.DotProduct(dir) ? a : b;
        }

        private static XYZ MaxEnd(Line l, XYZ dir)
        {
            XYZ a = l.GetEndPoint(0), b = l.GetEndPoint(1);
            return a.DotProduct(dir) > b.DotProduct(dir) ? a : b;
        }
    }
}
