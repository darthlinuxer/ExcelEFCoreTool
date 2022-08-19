namespace ExcelEFCore;

public partial class Worksheet
{
    public (int?, Element?) Find(Expression<Predicate<Element>> e)
    {
        try
        {
            var elementPredicate = e.Compile();
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            Element? rowElement = null;
            int row = 2;
            for (var line = 2; line <= GetLastRow(); line++)
            {
                rowElement = GetElement(line);
                var result = elementPredicate.Invoke(rowElement);
                if (result is true) break;
                row++;
            }
            Excel.Debug("{$a} {b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, rowElement);
            if (rowElement is not null) return (row, rowElement);
            return (null, null);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return (null, null);
        }
    }

    public IEnumerable<T> GetAll<T>() where T : class
    {
        Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
        Element? rowElement = null;
        var elements = new List<T>();
        int row = 1;
        for (var line = 2; line <= GetLastRow(); line++)
        {
            row++;
            rowElement = GetElement(row);
            if (rowElement is null) continue;
            T convertedRowElement = (rowElement.Item as T)!;
            elements.Add((rowElement.Item as T)!);
        }
        Excel.Debug("{$a} {b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, rowElement);
        return elements;
    }

    public bool Compare(Expression<Predicate<Element>> e, Element? target, Color color, bool compareId = false)
    {
        try
        {
            var elementPredicate = e.Compile();
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            Element? rowElement = null;
            int row = 2;
            for (var line = 2; line <= GetLastRow(); line++)
            {
                rowElement = GetElement(line);
                var result = elementPredicate.Invoke(rowElement);
                if (result is true) break;
                row++;
            }
            Excel.Debug("{$a} {b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, rowElement);
            if (rowElement is null) return false;
            var unMatchedProperties = CompareElements(rowElement, target, compareId);
            foreach (var unmatchedProperty in unMatchedProperties.Keys)
            {
                var col = this.HeaderProperties.ToList().IndexOf(this.HeaderProperties.FirstOrDefault(c => c.Name == unmatchedProperty)!) + 1;
                Cell.SetColor(RealWorksheet, row, col, color);
            }
            return unMatchedProperties?.Count() == 0;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return false;
        }
    }
}