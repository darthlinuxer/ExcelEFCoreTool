namespace ExcelEFCore;

public partial class Worksheet
{

    internal void Append(Element element)
    {
        try
        {
            var row = GetEmptyRow();
            Excel.Debug("{$a}:{b} row={c} {@d}", this, MethodBase.GetCurrentMethod()?.Name, row, element);
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
            Excel.Debug("{$a}:{b} row={c} {@d}", this, MethodBase.GetCurrentMethod()?.Name, row, element);
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
            Excel.Debug("{$a}:{b} {@d}", this, MethodBase.GetCurrentMethod()?.Name, element);
            var newRow = GetEmptyRow();
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

    public void AddRange(IEnumerable<object> items, string index="Id", CultureInfo? culture = null)
    {
        try
        {
            Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            foreach (var item in items){
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