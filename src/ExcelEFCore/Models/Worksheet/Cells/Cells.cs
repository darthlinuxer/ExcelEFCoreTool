namespace ExcelEFCore
{
    public class Cell
    {
        internal static int Row(ExcelAddress cell) => cell.Start.Row;
        internal static int Col(ExcelAddress cell) => cell.Start.Column;
        public static object? GetValue(ExcelWorksheet worksheet, int row, int col)
        {
            try
            {
                return worksheet.Cells[row, col].GetValue<object>();
            }
            catch (Exception ex)
            {
                Excel.Error(ex, "{a} {b}", MethodBase.GetCurrentMethod()?.Name, ex.Message);
                return null;
            }
        }

        public static void SetValue(ExcelWorksheet worksheet, int row, int col, object? value)
        {
            try
            {
                Excel.Debug("static {a}: worksheet={b} row={c} col={d} value={e}", MethodBase.GetCurrentMethod()?.Name, worksheet.Name, row, col, value);
                worksheet.Cells[row, col].Value = value;
            }
            catch (Exception ex)
            {
                Excel.Error(ex, "{a} {b}", MethodBase.GetCurrentMethod()?.Name, ex.Message);
            }
        }

        public static void SetColor(ExcelWorksheet worksheet, int row, int col, Color color)
        {
            try
            {
                Excel.Debug("static {a}: worksheet={b} row={c} col={d} color={e}", MethodBase.GetCurrentMethod()?.Name, worksheet.Name, row, col, color.ToKnownColor());
                worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(color);
            }
            catch (Exception ex)
            {
                Excel.Error(ex, "{a} {b}", MethodBase.GetCurrentMethod()?.Name, ex.Message);
            }
        }
    }
}