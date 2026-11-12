using Autodesk.Revit.DB;

namespace TD.DuctSplitter.Core
{
    /// <summary>Tự đóng các cảnh báo (Warning) phát sinh khi chia ống/đặt union để không hiện hộp thoại hàng loạt. Lỗi (Error) vẫn giữ nguyên.</summary>
    public class WarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            foreach (FailureMessageAccessor m in fa.GetFailureMessages())
            {
                if (m.GetSeverity() == FailureSeverity.Warning) fa.DeleteWarning(m);
            }
            return FailureProcessingResult.Continue;
        }

        public static void Attach(Transaction t)
        {
            FailureHandlingOptions o = t.GetFailureHandlingOptions();
            o.SetFailuresPreprocessor(new WarningSwallower());
            t.SetFailureHandlingOptions(o);
        }
    }
}
