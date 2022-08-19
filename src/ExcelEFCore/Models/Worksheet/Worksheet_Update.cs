namespace ExcelEFCore;

public partial class Worksheet
{
    public void UpdateWorksheetOnly(Element element)
    {
        try
        {
            Excel.Info("{$a}:{b} {@d}", this, MethodBase.GetCurrentMethod()?.Name, element);
            var (row, _) = Find(e => e.GetValue()!.Equals(element.GetValue()));
            if (row is not null) WriteToRow(row.Value, element);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    public void UpdateRange(IEnumerable<Element> elements)
    {
        try
        {
            Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            foreach (var element in elements) UpdateWorksheetOnly(element);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }

    }
}