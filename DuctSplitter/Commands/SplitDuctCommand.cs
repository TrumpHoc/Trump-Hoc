using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DialogResult = System.Windows.Forms.DialogResult;
using View = Autodesk.Revit.DB.View;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using TD.DuctSplitter.Core;
using TD.DuctSplitter.UI;

namespace TD.DuctSplitter.Commands
{
    /// <summary>Chia đốt ống gió hàng loạt theo chiều dài đốt điền trước + tự động dim Union trong view hiện hành (trừ 3D).</summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SplitDuctCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;
            Document doc = uidoc.Document;
            View view = uidoc.ActiveView;

            if (doc.IsFamilyDocument)
            {
                TaskDialog.Show("Chia đốt ống gió", "Lệnh chỉ chạy trong file dự án, không chạy trong Family Editor.");
                return Result.Cancelled;
            }

            SplitSettings settings = SplitSettings.Load();
            bool viewAllowsDims = UnionDimensionService.ViewSupportsDims(view, out string viewNote);
            List<string> dimTypes = UnionDimensionService.LinearDimensionTypes(doc).Select(t => t.Name).Distinct().ToList();

            bool dimThisRun;
            using (var form = new SplitDuctForm(settings, dimTypes, viewAllowsDims, viewNote))
            {
                if (form.ShowDialog() != DialogResult.OK) return Result.Cancelled;
                dimThisRun = form.DimThisRun;
            }
            settings.Save();

            IList<Duct> ducts = settings.Scope == ScopeMode.ActiveView
                ? DuctSelection.InView(doc, view)
                : DuctSelection.FromSelectionOrPick(uidoc, "Chọn các ống gió cần chia đốt, xong bấm Finish");
            if (ducts == null) return Result.Cancelled;
            if (ducts.Count == 0)
            {
                TaskDialog.Show("Chia đốt ống gió", "Không có ống gió (Duct) nào được chọn. Ống mềm (Flex Duct) không được xử lý.");
                return Result.Cancelled;
            }

            SplitReport report;
            int dimsCreated = 0;
            var dimErrors = new List<string>();

            using (var tg = new TransactionGroup(doc, "Chia đốt ống gió"))
            {
                tg.Start();

                using (var t = new Transaction(doc, "Chia đốt ống gió"))
                {
                    t.Start();
                    WarningSwallower.Attach(t);
                    report = new DuctSplitService(doc, settings).Run(ducts);
                    if (t.Commit() != TransactionStatus.Committed)
                    {
                        tg.RollBack();
                        TaskDialog.Show("Chia đốt ống gió", "Revit không cho phép hoàn tất thao tác chia ống – đã huỷ toàn bộ.");
                        return Result.Failed;
                    }
                }

                if (dimThisRun && report.Runs.Count > 0)
                {
                    var visible = new HashSet<ElementId>(DuctSelection.InView(doc, view).Select(d => d.Id));
                    using (var t = new Transaction(doc, "Dim Union"))
                    {
                        t.Start();
                        WarningSwallower.Attach(t);
                        var svc = new UnionDimensionService(doc, view, settings);
                        foreach (List<ElementId> run in report.Runs)
                        {
                            List<Duct> pieces = run.Select(id => doc.GetElement(id)).OfType<Duct>().ToList();
                            if (pieces.Count < 2) continue;
                            if (!pieces.All(p => visible.Contains(p.Id)))
                            {
                                dimErrors.Add($"Tuyến bắt đầu từ ống Id {pieces[0].Id}: không hiển thị đầy đủ trong view – bỏ qua dim.");
                                continue;
                            }
                            if (svc.DimensionRun(pieces, out string err) != null) dimsCreated++;
                            else dimErrors.Add($"Tuyến bắt đầu từ ống Id {pieces[0].Id}: {err}");
                        }
                        t.Commit();
                    }
                }

                tg.Assimilate();
            }

            var sb = new StringBuilder();
            sb.AppendLine($"Số ống xét: {report.DuctsProcessed}");
            sb.AppendLine($"Số ống đã chia: {report.DuctsSplit}");
            sb.AppendLine($"Số ống không cần chia (ngắn hơn 1 đốt + đốt cuối tối thiểu): {report.DuctsSkippedShort}");
            sb.AppendLine($"Số mối chia: {report.BreaksCreated}  (Union tạo được: {report.UnionsCreated}, nối không Union: {report.JointsWithoutUnion})");
            if (dimThisRun) sb.AppendLine($"Số chuỗi dim đã tạo trong view \"{view.Name}\": {dimsCreated}");
            else if (!viewAllowsDims) sb.AppendLine($"Không tạo dim: {viewNote}");

            var notes = report.Messages.Concat(dimErrors).ToList();
            var td = new TaskDialog("Chia đốt ống gió")
            {
                MainInstruction = "Hoàn tất",
                MainContent = sb.ToString()
            };
            if (notes.Count > 0)
                td.ExpandedContent = string.Join(Environment.NewLine, notes.Take(50)) +
                                     (notes.Count > 50 ? $"{Environment.NewLine}... và {notes.Count - 50} ghi chú khác" : "");
            td.Show();
            return Result.Succeeded;
        }
    }
}
