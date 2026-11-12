using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace TD.DuctSplitter.Core
{
    /// <summary>
    /// Gom các ống đã chọn thành từng tuyến thẳng: các ống thẳng hàng nối với nhau qua Union
    /// (hoặc nối trực tiếp ống–ống). Dùng cho lệnh "Dim Union" độc lập.
    /// </summary>
    public static class RunBuilder
    {
        public static List<List<Duct>> Build(IList<Duct> ducts)
        {
            var byId = ducts.ToDictionary(d => d.Id);
            var adj = ducts.ToDictionary(d => d.Id, d => new List<ElementId>());

            foreach (Duct d in ducts)
            {
                Line ld = DuctHelper.GetLine(d);
                foreach (Connector c in DuctHelper.EndConnectors(d.ConnectorManager))
                {
                    foreach (Element owner in DuctHelper.ConnectedOwners(c))
                    {
                        var neighbours = new List<Element>();
                        if (owner is Duct) neighbours.Add(owner);
                        else if (DuctHelper.IsUnion(owner))
                        {
                            var fi = (FamilyInstance)owner;
                            foreach (Connector cu in DuctHelper.EndConnectors(fi.MEPModel.ConnectorManager))
                                neighbours.AddRange(DuctHelper.ConnectedOwners(cu).Where(o => o.Id != d.Id));
                        }

                        foreach (Element n in neighbours)
                        {
                            if (!byId.TryGetValue(n.Id, out Duct nd)) continue;
                            if (!DuctHelper.IsCollinear(ld, DuctHelper.GetLine(nd))) continue;
                            if (!adj[d.Id].Contains(nd.Id)) adj[d.Id].Add(nd.Id);
                            if (!adj[nd.Id].Contains(d.Id)) adj[nd.Id].Add(d.Id);
                        }
                    }
                }
            }

            var visited = new HashSet<ElementId>();
            var runs = new List<List<Duct>>();
            foreach (Duct d in ducts)
            {
                if (visited.Contains(d.Id)) continue;
                var comp = new List<Duct>();
                var stack = new Stack<ElementId>();
                stack.Push(d.Id);
                visited.Add(d.Id);
                while (stack.Count > 0)
                {
                    ElementId id = stack.Pop();
                    comp.Add(byId[id]);
                    foreach (ElementId n in adj[id])
                        if (visited.Add(n)) stack.Push(n);
                }

                XYZ dir = DuctHelper.GetLine(comp[0]).Direction;
                runs.Add(comp.OrderBy(x => DuctHelper.GetLine(x).Evaluate(0.5, true).DotProduct(dir)).ToList());
            }
            return runs;
        }
    }
}
