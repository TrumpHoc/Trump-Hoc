using System;
using System.Reflection;
using Autodesk.Revit.UI;

namespace TD.DuctSplitter
{
    /// <summary>Tạo tab ribbon "TD MEP" > panel "Ống gió" với 2 nút lệnh.</summary>
    public class App : IExternalApplication
    {
        public const string TabName = "TD MEP";

        public Result OnStartup(UIControlledApplication app)
        {
            try { app.CreateRibbonTab(TabName); } catch { /* tab đã tồn tại (add-in khác cùng tab) */ }

            RibbonPanel panel = app.CreateRibbonPanel(TabName, "Ống gió");
            string asm = Assembly.GetExecutingAssembly().Location;

            var split = new PushButtonData("TD_SplitDuct", "Chia đốt\nống gió", asm, typeof(Commands.SplitDuctCommand).FullName)
            {
                ToolTip = "Chia đốt ống gió hàng loạt theo chiều dài đốt nhập trước, nối bằng Union (Routing Preferences) " +
                          "và tự động dim các Union trong mặt bằng đang mở (không dim trong 3D view).",
                LargeImage = Icons.Split(32),
                Image = Icons.Split(16)
            };
            var dim = new PushButtonData("TD_DimUnions", "Dim\nUnion", asm, typeof(Commands.DimUnionsCommand).FullName)
            {
                ToolTip = "Dim các Union trên những tuyến ống gió đã chia sẵn, trong view mặt bằng/mặt cắt hiện hành.",
                LargeImage = Icons.Dim(32),
                Image = Icons.Dim(16)
            };

            panel.AddItem(split);
            panel.AddItem(dim);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;
    }
}
