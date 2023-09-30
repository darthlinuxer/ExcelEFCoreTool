namespace ExcelEFCore;

public partial class Worksheet
{

    internal void Append(Element element)
    {
        try
        {
            var row = GetEmptyRow();
            Excel.Debug("{$a}:{b} adding element:{c} key value:{d} to row {e}", this, MethodBase.GetCurrentMethod()?.Name, element.GetType().Name, element.GetValue(), row);
            WriteToRow(row, element);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}: {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }
    internal void Add(int row, Element element)
    {
        try
        {
            Excel.Debug("{$a}:{b} adding element:{c} key value:{d} to row {e}", this, MethodBase.GetCurrentMethod()?.Name, element.GetType().Name, element.GetValue(), row);
            WriteToRow(row, element);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}: {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    public void Add(Element element)
    {
        try
        {
            var newRow = GetEmptyRow();
            Excel.Debug("{$a}:{b} adding element:{c} key value:{d} to row {e}", this, MethodBase.GetCurrentMethod()?.Name, element.GetType().Name, element.GetValue(), newRow);
            Add(newRow, element);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    public void AddRange(IEnumerable<Element> elements)
    {
        try
        {
            Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            foreach (var element in elements) Add(element);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    public void AddRange(IEnumerable<object> items, string index = "Id", CultureInfo? culture = null)
    {
        try
        {
            Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            foreach (var item in items)
            {
                var element = Element.Factory(item, index, culture);
                Add(element);
            }
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }
}