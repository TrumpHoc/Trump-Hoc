using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using TD.DuctSplitter.Core;

namespace TD.DuctSplitter.Commands
{
    /// <summary>Dim các Union trên những tuyến ống gió đã chia sẵn (chọn ống trong view hiện hành, trừ 3D).</summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DimUnionsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;
            Document doc = uidoc.Document;
            View view = uidoc.ActiveView;

            if (!UnionDimensionService.ViewSupportsDims(view, out string reason))
            {
                TaskDialog.Show("Dim Union", "Không tạo dim: " + reason + Environment.NewLine + "Hãy mở mặt bằng (hoặc mặt cắt) rồi chạy lại.");
                return Result.Cancelled;
            }

            IList<Duct> ducts = DuctSelection.FromSelectionOrPick(uidoc, "Chọn các đoạn ống gió (đã chia) cần dim Union, xong bấm Finish");
            if (ducts == null) return Result.Cancelled;
            if (ducts.Count == 0)
            {
                TaskDialog.Show("Dim Union", "Không có ống gió nào được chọn.");
                return Result.Cancelled;
            }

            SplitSettings settings = SplitSettings.Load();
            List<List<Duct>> runs = RunBuilder.Build(ducts).Where(r => r.Count >= 2).ToList();
            int created = 0;
            var errors = new List<string>();

            using (var t = new Transaction(doc, "Dim Union"))
            {
                t.Start();
                WarningSwallower.Attach(t);
                var svc = new UnionDimensionService(doc, view, settings);
                foreach (List<Duct> run in runs)
                {
                    if (svc.DimensionRun(run, out string err) != null) created++;
                    else errors.Add($"Tuyến bắt đầu từ ống Id {run[0].Id}: {err}");
                }
                t.Commit();
            }

            var td = new TaskDialog("Dim Union")
            {
                MainInstruction = "Hoàn tất",
                MainContent = $"Số tuyến có Union: {runs.Count}{Environment.NewLine}Số chuỗi dim đã tạo trong view \"{view.Name}\": {created}" +
                              $"{Environment.NewLine}(Kiểu dim / khoảng cách lấy theo lần chạy lệnh Chia đốt gần nhất)"
            };
            if (errors.Count > 0) td.ExpandedContent = string.Join(Environment.NewLine, errors.Take(50));
            td.Show();
            return Result.Succeeded;
        }
    }
}
