using NpoiSheet = NPOI.SS.UserModel.ISheet;

namespace ExcelReport.Driver.NPOI.Extends
{
    internal static class SheetExtend
    {
        public static Sheet GetAdapter(this NpoiSheet sheet)
        {
            if (null == sheet)
            {
                return null;
            }
            sheet.ForceFormulaRecalculation = true;//强制刷新函数
            return new Sheet(sheet);
        }
    }
}