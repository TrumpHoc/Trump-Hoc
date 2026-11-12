using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace TD.DuctSplitter.Core
{
    public class DuctFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is Duct;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    public static class DuctSelection
    {
        /// <summary>Lấy ống gió (Duct, không gồm ống mềm) từ selection hiện tại; nếu trống thì cho pick. Trả về null nếu người dùng huỷ.</summary>
        public static IList<Duct> FromSelectionOrPick(UIDocument uidoc, string prompt)
        {
            Document doc = uidoc.Document;
            List<Duct> ducts = uidoc.Selection.GetElementIds()
                .Select(id => doc.GetElement(id))
                .OfType<Duct>()
                .ToList();
            if (ducts.Count > 0) return ducts;

            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(ObjectType.Element, new DuctFilter(), prompt);
                return refs.Select(r => doc.GetElement(r)).OfType<Duct>().ToList();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        public static IList<Duct> InView(Document doc, View view)
        {
            return new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Duct))
                .Cast<Duct>()
                .ToList();
        }
    }
}
