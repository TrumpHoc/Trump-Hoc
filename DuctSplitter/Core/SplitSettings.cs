using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace TD.DuctSplitter.Core
{
    /// <summary>Đầu ống bắt đầu đo chiều dài đốt.</summary>
    public enum StartSide
    {
        /// <summary>Tự động: ưu tiên đầu đang nối với fitting/thiết bị; nếu cả 2 (hoặc không đầu nào) nối thì lấy điểm Start của đường ống.</summary>
        Auto = 0,
        /// <summary>Điểm Start của LocationCurve.</summary>
        CurveStart = 1,
        /// <summary>Điểm End của LocationCurve.</summary>
        CurveEnd = 2
    }

    /// <summary>Phạm vi chọn ống.</summary>
    public enum ScopeMode
    {
        /// <summary>Các ống đang chọn (nếu chưa chọn sẽ yêu cầu pick).</summary>
        Selection = 0,
        /// <summary>Tất cả ống gió nhìn thấy trong view hiện hành.</summary>
        ActiveView = 1
    }

    /// <summary>
    /// Thông số chia đốt. Mọi chiều dài lưu theo mm (đơn vị người dùng),
    /// chỉ đổi sang feet (đơn vị nội bộ Revit) khi tính toán.
    /// </summary>
    public class SplitSettings
    {
        public double RectLengthMm { get; set; } = 1200;   // ống chữ nhật
        public double RoundLengthMm { get; set; } = 1200;  // ống tròn
        public double OvalLengthMm { get; set; } = 1200;   // ống oval
        public double MinTailMm { get; set; } = 100;       // đốt cuối ngắn hơn giá trị này thì gộp vào đốt trước
        public StartSide StartSide { get; set; } = StartSide.Auto;
        public ScopeMode Scope { get; set; } = ScopeMode.Selection;
        public bool CreateDimensions { get; set; } = true;
        public bool DimIncludeEnds { get; set; } = true;   // dim cả 2 đầu ống (tổng chuỗi)
        public double DimOffsetMm { get; set; } = 8;       // khoảng cách đường dim tính từ mép ống, mm trên giấy (nhân tỉ lệ view)
        public string DimensionTypeName { get; set; } = "";

        public const double MmPerFoot = 304.8;
        public static double MmToFt(double mm) => mm / MmPerFoot;
        public static double FtToMm(double ft) => ft * MmPerFoot;

        private static string FilePath
        {
            get
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TD.DuctSplitter");
                return Path.Combine(dir, "settings.txt");
            }
        }

        public static SplitSettings Load()
        {
            var s = new SplitSettings();
            try
            {
                if (!File.Exists(FilePath)) return s;
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (string line in File.ReadAllLines(FilePath))
                {
                    int i = line.IndexOf('=');
                    if (i > 0) map[line.Substring(0, i).Trim()] = line.Substring(i + 1).Trim();
                }
                s.RectLengthMm = GetD(map, nameof(RectLengthMm), s.RectLengthMm);
                s.RoundLengthMm = GetD(map, nameof(RoundLengthMm), s.RoundLengthMm);
                s.OvalLengthMm = GetD(map, nameof(OvalLengthMm), s.OvalLengthMm);
                s.MinTailMm = GetD(map, nameof(MinTailMm), s.MinTailMm);
                s.DimOffsetMm = GetD(map, nameof(DimOffsetMm), s.DimOffsetMm);
                s.StartSide = (StartSide)(int)GetD(map, nameof(StartSide), (int)s.StartSide);
                s.Scope = (ScopeMode)(int)GetD(map, nameof(Scope), (int)s.Scope);
                s.CreateDimensions = GetD(map, nameof(CreateDimensions), 1) > 0.5;
                s.DimIncludeEnds = GetD(map, nameof(DimIncludeEnds), 1) > 0.5;
                if (map.TryGetValue(nameof(DimensionTypeName), out string dt)) s.DimensionTypeName = dt;
            }
            catch
            {
                // Lỗi đọc file cấu hình -> dùng mặc định
            }
            return s;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                var ci = CultureInfo.InvariantCulture;
                File.WriteAllLines(FilePath, new[]
                {
                    $"{nameof(RectLengthMm)}={RectLengthMm.ToString(ci)}",
                    $"{nameof(RoundLengthMm)}={RoundLengthMm.ToString(ci)}",
                    $"{nameof(OvalLengthMm)}={OvalLengthMm.ToString(ci)}",
                    $"{nameof(MinTailMm)}={MinTailMm.ToString(ci)}",
                    $"{nameof(DimOffsetMm)}={DimOffsetMm.ToString(ci)}",
                    $"{nameof(StartSide)}={(int)StartSide}",
                    $"{nameof(Scope)}={(int)Scope}",
                    $"{nameof(CreateDimensions)}={(CreateDimensions ? 1 : 0)}",
                    $"{nameof(DimIncludeEnds)}={(DimIncludeEnds ? 1 : 0)}",
                    $"{nameof(DimensionTypeName)}={DimensionTypeName}"
                });
            }
            catch
            {
                // Không lưu được cấu hình thì bỏ qua, không ảnh hưởng lệnh
            }
        }

        private static double GetD(Dictionary<string, string> map, string key, double def)
        {
            return map.TryGetValue(key, out string v) &&
                   double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) ? d : def;
        }
    }
}
