using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace TD.DuctSplitter
{
    /// <summary>Icon ribbon vẽ bằng code (không cần file ảnh kèm theo).</summary>
    internal static class Icons
    {
        private static readonly Brush Duct = new SolidColorBrush(Color.FromRgb(0x3A, 0x7B, 0xD5));
        private static readonly Brush Accent = new SolidColorBrush(Color.FromRgb(0xE0, 0x4F, 0x2F));
        private static readonly Brush Ink = new SolidColorBrush(Color.FromRgb(0x30, 0x30, 0x30));

        public static ImageSource Split(int size) => Render(size, (dc, s) =>
        {
            double h = s * 0.34, y = (s - h) / 2;
            dc.DrawRectangle(Duct, null, new Rect(s * 0.04, y, s * 0.92, h));
            var pen = new Pen(Accent, s / 12.0);
            foreach (double x in new[] { s * 0.35, s * 0.65 })
                dc.DrawLine(pen, new Point(x, y - s * 0.12), new Point(x, y + h + s * 0.12));
        });

        public static ImageSource Dim(int size) => Render(size, (dc, s) =>
        {
            double h = s * 0.28, y = s * 0.55;
            dc.DrawRectangle(Duct, null, new Rect(s * 0.04, y, s * 0.92, h));
            var pen = new Pen(Ink, s / 20.0);
            double dy = s * 0.25;
            dc.DrawLine(pen, new Point(s * 0.06, dy), new Point(s * 0.94, dy));
            foreach (double x in new[] { s * 0.06, s * 0.5, s * 0.94 })
            {
                dc.DrawLine(pen, new Point(x, dy - s * 0.12), new Point(x, y));
                dc.DrawLine(new Pen(Accent, s / 14.0), new Point(x - s * 0.06, dy + s * 0.06), new Point(x + s * 0.06, dy - s * 0.06));
            }
        });

        private static ImageSource Render(int size, System.Action<DrawingContext, double> draw)
        {
            var dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen()) draw(dc, size);
            var bmp = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(dv);
            bmp.Freeze();
            return bmp;
        }
    }
}
