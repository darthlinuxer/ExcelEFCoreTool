namespace ExcelEFCore;

public partial class Worksheet
{
    public string Name { get; init; }
    private CultureInfo Culture = CultureInfo.CurrentCulture;
    private Color uIFontHeaderColor = Color.Black;
    private Color uIHeaderColor = Color.LightBlue;
    private ExcelWorksheet RealWorksheet { get; init; }
    private PropertyInfo ContextDbSet { get; init; }
    private Type ElementType { get; init; }
    private string KeyName { get; init; }
    private int KeyCol { get; init; }
    private PropertyInfo KeyProp { get; init; }
    private IEnumerable<PropertyInfo> HeaderProperties { get; init; }
    internal Color UIUpdateColor { get; set; } = Color.Yellow;
    internal Color UIDeleteColor { get; set; } = Color.Red;
    internal Color UIAddColor { get; set; } = Color.Blue;
    internal Color UISuccessColor { get; set; } = Color.LightGreen;
    internal Color UIHeaderColor
    {
        get => uIHeaderColor;
        set
        {
            uIHeaderColor = value;
            StyleHeaders();
        }
    }

    internal Color UIFontHeaderColor
    {
        get => uIFontHeaderColor;
        set
        {
            uIFontHeaderColor = value;
            StyleHeaders();
        }
    }


    private Worksheet(PropertyInfo dbSet, ExcelWorksheet worksheet, bool OverwriteInvalidHeaders, string keyName, CultureInfo? culture = null)
    {
        try
        {
            if (culture is not null) Culture = culture;
            Excel.Info("{$a}:{b} DbSet={c} Worksheet={d}", this, MethodBase.GetCurrentMethod()?.Name, dbSet.Name, worksheet.Name);
            RealWorksheet = worksheet;
            if (RealWorksheet.Name != dbSet.Name) throw new Exception("Worksheet NAME do not match any DbSet NAME in context!");
            ContextDbSet = dbSet;
            Name = RealWorksheet.Name;
            ElementType = dbSet.PropertyType.GenericTypeArguments.Single().UnderlyingSystemType;
            Excel.Debug("{$a} ElementType:{b}", this, ElementType.Name);
            HeaderProperties = ElementType.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly).Where(c => c.CustomAttributes.All(a => a.AttributeType != typeof(EpplusIgnore)));
            Excel.Debug("{$a} HeaderProperties from ElementType: {@b}", this, HeaderProperties.Select(c => c.Name));
            var isKeyValid = HeaderProperties.Select(c => c.Name).Contains(keyName);
            Excel.Debug("{$a} Key={b} is valid ? {c}", this, keyName, isKeyValid);
            if (!isKeyValid) throw new Exception($"{keyName} is not a valid property of {ElementType.Name}");
            KeyName = keyName;
            KeyCol = HeaderProperties.Select(c => c.Name).ToList().IndexOf(KeyName) + 1;
            KeyProp = HeaderProperties.Single(c => c.Name == keyName);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }

    }

    internal static Worksheet Factory(PropertyInfo dbSet, ExcelWorksheet worksheet, bool overwriteInvalidHeaders = true, string key = "Id", CultureInfo? culture = null)
    {
        try
        {
            if (dbSet.Name != worksheet.Name) throw new Exception($"Worksheet {worksheet.Name} is not a DbSet of Context");
            var realWkSheet = new Worksheet(dbSet, worksheet, overwriteInvalidHeaders, key, culture);
            var validHeaders = realWkSheet.ValidateHeaders();
            Excel.Debug("{$a} ValidHeaders={b}", MethodBase.GetCurrentMethod()?.Name, validHeaders);
            if (!validHeaders && overwriteInvalidHeaders) realWkSheet.CreateHeaders();
            if (validHeaders && !overwriteInvalidHeaders) throw new Exception($"{worksheet.Name} have INVALID Headers!");
            return realWkSheet;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a} {b}", MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    internal IEnumerable<Error> ValidateWorksheetParameters()
    {
        try
        {
            var errors = new List<Error>();
            var sheetName = RealWorksheet.Name;
            if (ContextDbSet.Name != sheetName) errors.Add(new Error() { Message = $"SheetName does not exist in Context! Rename to {ContextDbSet.Name}" });
            if (InvalidHeaders()?.Count() > 0) errors.Add(new Error() { Message = "Invalid Headers", item = InvalidHeaders() });
            return errors;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

}