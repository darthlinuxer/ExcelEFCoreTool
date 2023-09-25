namespace ExcelEFCore;

public partial class Element
{
    internal bool SetValue(object value)
    {
        return SetValue(value, Key);
    }
    internal bool SetValue(object? value, PropertyInfo property)
    {
        try
        {
            Excel.Info("{$a}:{b} property {c}={d}", this, MethodBase.GetCurrentMethod()?.Name, property.Name, value);
            Type t = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            if (value?.GetType().Name == "String" && t?.Name == "DateOnly")
            {
                DateOnly dateResult;
                string? stringValue = Convert.ChangeType(value, typeof(string)) as string;
                DateOnly.TryParse(stringValue, Culture, DateTimeStyles.None, out dateResult);
                value = dateResult;
            }

            if (value?.GetType().Name == "Double" && t?.Name == "DateTime")
            {
                DateTime datetimeResult;
                string? stringValue = Convert.ChangeType(value, typeof(string)) as string;
                DateTime.TryParse(stringValue, Culture, DateTimeStyles.None, out datetimeResult);
                value = datetimeResult;
            }

            if (value is null && property.PropertyType != typeof(Nullable)) throw new Exception($"Can't set null the non-nullable property {property.Name} of type {property.GetType()}");

            var convertedValue = Convert.ChangeType(value, t!, Culture);
            Excel.Debug("{$a}:{b} PropertyType={c}, ConvertedValue={d}", this, MethodBase.GetCurrentMethod()?.Name, t!.GetType(), convertedValue);
            property.SetValue(Item, convertedValue);
            return true;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return false;
        }
    }
}