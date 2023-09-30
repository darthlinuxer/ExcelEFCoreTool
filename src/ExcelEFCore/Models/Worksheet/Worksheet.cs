namespace ExcelEFCore;

public partial class Worksheet
{
    private int GetLastRow() => RealWorksheet.Dimension.Rows;
    private int GetEmptyRow() => GetLastRow() + 1;
    private void DeleteRow(int row) => RealWorksheet.DeleteRow(row);
    private bool WriteHeaders(IEnumerable<string> values)
    {
        try
        {
            var col = 1;
            foreach (var value in values)
            {
                Cell.SetValue(RealWorksheet, 1, col, value);
                col++;
            }
            return true;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            return false;
        }

    }
    private bool WriteToRow(int row, Element element)
    {
        try
        {
            Excel.Info("{$a}:{b} row={c} element key value:{d}", this, MethodBase.GetCurrentMethod()?.Name, row, element.GetValue());
            if (!ValidateHeaders(element)) return false;
            foreach (var header in HeaderProperties)
            {
                var value = element.GetValue(header);
                var col = HeaderProperties.Select(c => c.Name).ToList().IndexOf(header.Name) + 1;
                Excel.Debug("{$a}:{b} header={c} row={d} col={e} value={f}", this, MethodBase.GetCurrentMethod()?.Name, header.Name, row, col, value);
                Cell.SetValue(RealWorksheet, row, col, value);
            }
            return true;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return false;
        }
    }

    private Dictionary<string, object?> CompareElements(Element source, Element? target, bool compareId = false)
    {
        try
        {
            Excel.Info("{$a}:{b} source={@c} target={@d} ", this, MethodBase.GetCurrentMethod()?.Name, source.Item, target?.Item);
            return source.UnMatchedProperties(target, compareId);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return new Dictionary<string, object?>();
        }
    }

    private Element GetElement(int row)
    {
        try
        {
            Excel.Info("{$a}:{b} row={c}", this, MethodBase.GetCurrentMethod()?.Name, row);
            var contextObject = Activator.CreateInstance(ElementType);
            if (contextObject is null) throw new Exception($"{ElementType.Name} not found in Assembly!");
            var element = Element.Factory(contextObject, KeyName, Culture);
            foreach (var property in HeaderProperties)
            {
                var col = HeaderProperties.Select(c => c.Name).ToList().IndexOf(property.Name) + 1;
                var value = Cell.GetValue(RealWorksheet, row, col);
                element.SetValue(value, property);
            }
            Excel.Debug("{$a}:{b} element key value:{c}", this, MethodBase.GetCurrentMethod()?.Name, element?.GetValue());
            return element;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    private IEnumerable<int> RowsToUpdate()
    {
        var rows = RealWorksheet.Dimension.Rows;
        var cols = RealWorksheet.Dimension.Columns;
        var coloredCells = RealWorksheet.Cells[2, 1, rows, cols].Where(c =>
        {
            var color = ColorTranslator.FromHtml(c.Style.Fill.BackgroundColor.LookupColor());
            return color == UIUpdateColor;
        });
        var filteredCells = coloredCells.Select(c => c.EntireRow.StartRow).Distinct().Where(row => row != 1);
        return filteredCells;
    }
    private IEnumerable<int>? RowsToAdd()
    {
        var rows = RealWorksheet.Dimension.Rows;
        var cols = RealWorksheet.Dimension.Columns;
        var coloredCells = RealWorksheet.Cells[2, 1, rows, cols].Where(c =>
        {
            var color = ColorTranslator.FromHtml(c.Style.Fill.BackgroundColor.LookupColor());
            return color == UIAddColor;
        });
        var filteredCells = coloredCells.Select(c => c.EntireRow.StartRow).Distinct().Where(row => row != 1);
        return filteredCells;
    }

    private IEnumerable<int>? RowsToDelete()
    {
        var rows = RealWorksheet.Dimension.Rows;
        var cols = RealWorksheet.Dimension.Columns;
        var coloredCells = RealWorksheet.Cells[2, 1, rows, cols].Where(c =>
        {
            var color = ColorTranslator.FromHtml(c.Style.Fill.BackgroundColor.LookupColor());
            var result = color == UIDeleteColor;
            return result;
        });
        var filteredCells = coloredCells.Select(d => d.EntireRow.StartRow).Distinct().Where(row => row != 1);
        return filteredCells;
    }

    public void ColorRow(int row, Color? color = null)
    {
        var col = RealWorksheet.Dimension.Columns;
        RealWorksheet.Cells[row, 1, row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
        if (color is null)
            RealWorksheet.Cells[row, 1, row, col].Style.Fill.BackgroundColor.SetColor(Color.Transparent);
        else
            RealWorksheet.Cells[row, 1, row, col].Style.Fill.BackgroundColor.SetColor(color.Value);
    }

    public void ColorRow(Element element, Color? color = null)
    {
        var (row, _) = Find(e => e.GetValue() == element.GetValue());
        if (row is null) return;
        ColorRow(row.Value, color);
    }

    internal void ColorRow(IEnumerable<Element> elements)
    {
        foreach (var element in elements) ColorRow(element);
    }

}