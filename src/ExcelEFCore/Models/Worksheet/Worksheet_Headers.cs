namespace ExcelEFCore;

public partial class Worksheet
{
    private bool ValidateHeaders()
    {

        Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
        var headers = GetHeaders();
        return HeaderProperties.Select(c => c.Name).All(c => headers.Contains(c));
    }

    private bool ValidateHeaders(Element element){
        Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
        var elementProperties = element.Properties.Select(c=>c.Name);
        return HeaderProperties.Select(c=>c.Name).All(c=> elementProperties.Contains(c));
    }

    private IEnumerable<string> InvalidHeaders() => HeaderProperties.Select(c => c.Name).Except(GetHeaders());
    private void CreateHeaders()
    {
        try
        {
            var headers = HeaderProperties.Select(c => c.Name);
            Excel.Debug("{$a} {b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, headers);
            WriteHeaders(headers);
            StyleHeaders();
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }
    private void StyleHeaders()
    {
        try
        {
            Excel.Debug("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            RealWorksheet.Cells["1:1"].AutoFitColumns();
            RealWorksheet.Cells["1:1"].Style.Font.Bold = true;
            RealWorksheet.Cells["1:1"].Style.Font.Color.SetColor(UIFontHeaderColor);
            RealWorksheet.Cells["1:1"].Style.Fill.PatternType = ExcelFillStyle.None;
            RealWorksheet.Cells["1:1"].Style.Fill.SetBackground(UIHeaderColor);
            RealWorksheet.Cells["1:1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    private List<string> GetHeaders()
    {
        Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
        var headers = new List<string>();
        try
        {
            var hdrRange = this.RealWorksheet.Dimension.Columns;
            var col = 1;
            while (col <= hdrRange)
            {
                var hdrValue = Cell.GetValue(RealWorksheet, 1, col) as string;
                if (hdrValue is not null) headers.Add(hdrValue!);
                col++;
            }

            Excel.Debug("{$a}:{b} worksheet:{c} headers={@d}", this, MethodBase.GetCurrentMethod()?.Name, RealWorksheet.Name, headers);
            return headers;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return headers;
        }
    }

}