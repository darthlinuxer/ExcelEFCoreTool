namespace ExcelEFCore;

public partial class Excel
{
    public IEnumerable<T>? GetElementsFromWorksheet<T>(string name) where T : class
    {
        try
        {
            var worksheet = Worksheets.FirstOrDefault(c => c.Name == name);
            return worksheet?.GetAll<T>();
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return null;
        }
    }
}