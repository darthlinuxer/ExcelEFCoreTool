namespace ExcelEFCore;

public partial class Worksheet
{
    public void ClearAll(bool propagateClearToContext = true)
    {
        try
        {
            Excel.Info("{$a}:{b} propagateClearToContext:{c}", this, MethodBase.GetCurrentMethod()?.Name, propagateClearToContext);
            this.RealWorksheet.ClearFormulas();
            var dimension = this.RealWorksheet.Dimension;
            this.RealWorksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            this.RealWorksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.Transparent);
            this.RealWorksheet.Cells.Clear();
            CreateHeaders();
            if (propagateClearToContext)
            {
                var result = ClearAllEvent?.Invoke(ContextDbSet);
            }
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    public void Remove(Element element)
    {
        try
        {
            Excel.Debug("{$a}:{b} {@d}", this, MethodBase.GetCurrentMethod()?.Name, element);
            var (row, _) = Find(e => e.GetValue() == element.GetValue());
            if (row is not null) DeleteRow(row.Value);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    public void Remove(object keyValue)
    {
        try
        {
            Excel.Debug("{$a}:{b} index={d}", this, MethodBase.GetCurrentMethod()?.Name, keyValue);
            var (row, _) = Find(e => e.GetValue() == keyValue);
            if (row is not null) DeleteRow(row.Value);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    public void RemoveRange(Func<IEnumerable<Element>> p)
    {
        try
        {
            Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            var elements = p.Invoke();
            foreach (var element in elements) Remove(element);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }


}