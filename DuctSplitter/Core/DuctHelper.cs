using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace TD.DuctSplitter.Core
{
    /// <summary>Các hàm tiện ích hình học / kết nối cho ống gió.</summary>
    public static class DuctHelper
    {
        /// <summary>Dung sai trùng điểm (feet) ~ 1.5 mm.</summary>
        public const double PointTol = 0.005;

        public static Line GetLine(Duct duct)
        {
            return (duct?.Location as LocationCurve)?.Curve as Line;
        }

        public static IEnumerable<Connector> EndConnectors(ConnectorManager cm)
        {
            if (cm == null) yield break;
            foreach (Connector c in cm.Connectors)
            {
                if (c.ConnectorType == ConnectorType.End) yield return c;
            }
        }

        public static Connector ConnectorAt(Duct duct, XYZ pt, double tol = PointTol)
        {
            return EndConnectors(duct.ConnectorManager)
                .OrderBy(c => c.Origin.DistanceTo(pt))
                .FirstOrDefault(c => c.Origin.DistanceTo(pt) < tol);
        }

        public static ConnectorProfileType Shape(Duct duct)
        {
            try { return duct.DuctType.Shape; }
            catch
            {
                Connector c = EndConnectors(duct.ConnectorManager).FirstOrDefault();
                return c?.Shape ?? ConnectorProfileType.Rectangular;
            }
        }

        /// <summary>Kích thước ống theo phương ngang trong mặt bằng (feet) – dùng để đặt đường dim ra ngoài mép ống.</summary>
        public static double PlanWidth(Duct duct)
        {
            Parameter w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            if (w != null && w.HasValue && w.AsDouble() > 0) return w.AsDouble();
            Parameter d = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            if (d != null && d.HasValue && d.AsDouble() > 0) return d.AsDouble();
            return 0;
        }

        /// <summary>Đầu ống có nối với phần tử khác (fitting, thiết bị, ống...) hay không.</summary>
        public static bool IsEndConnected(Duct duct, XYZ end)
        {
            Connector c = ConnectorAt(duct, end);
            return c != null && c.IsConnected;
        }

        public static bool IsUnion(Element e)
        {
            if (e is FamilyInstance fi && fi.MEPModel is MechanicalFitting mf)
                return mf.PartType == PartType.Union;
            return false;
        }

        /// <summary>Tìm các phần tử nối trực tiếp với connector (không tính chính owner).</summary>
        public static IEnumerable<Element> ConnectedOwners(Connector c)
        {
            if (c == null || !c.IsConnected) yield break;
            foreach (Connector r in c.AllRefs)
            {
                if (r.Owner == null || r.Owner.Id == c.Owner.Id) continue;
                if (r.ConnectorType != ConnectorType.End && r.ConnectorType != ConnectorType.Curve) continue;
                yield return r.Owner;
            }
        }

        /// <summary>Union nằm giữa 2 ống a, b (nếu có).</summary>
        public static FamilyInstance UnionBetween(Duct a, Duct b)
        {
            foreach (Connector ca in EndConnectors(a.ConnectorManager))
            {
                foreach (Element owner in ConnectedOwners(ca))
                {
                    if (!IsUnion(owner)) continue;
                    var fi = (FamilyInstance)owner;
                    foreach (Connector cu in EndConnectors(fi.MEPModel.ConnectorManager))
                    {
                        if (ConnectedOwners(cu).Any(o => o.Id == b.Id)) return fi;
                    }
                }
            }
            return null;
        }

        public static bool IsParallel(XYZ a, XYZ b, double tol = 1e-3)
        {
            return a.Normalize().CrossProduct(b.Normalize()).GetLength() < tol;
        }

        /// <summary>2 ống thẳng hàng (song song và cùng nằm trên một đường thẳng).</summary>
        public static bool IsCollinear(Line a, Line b, double tol = 0.01)
        {
            if (a == null || b == null) return false;
            if (!IsParallel(a.Direction, b.Direction)) return false;
            XYZ v = b.GetEndPoint(0) - a.GetEndPoint(0);
            XYZ perp = v - a.Direction.Multiply(v.DotProduct(a.Direction));
            return perp.GetLength() < tol;
        }
    }
}
