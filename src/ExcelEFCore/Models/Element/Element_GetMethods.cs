namespace ExcelEFCore;

public partial class Element
{
    private IEnumerable<PropertyInfo> GetProperties(object item)
    {
        try
        {
            var properties = item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(c => c.CustomAttributes.All(a => a.AttributeType != typeof(EpplusIgnore)));
            if (properties is null) throw new Exception("Item sent must have public properties");
            Excel.Debug("{$a}:{b} properties:{@c}", this, MethodBase.GetCurrentMethod()?.Name, properties?.Select(c => new { name = c.Name }));
            return properties!;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }
    internal PropertyInfo? Property(string name)
    {
        Excel.Debug("{$a}:{b}({name})", this, MethodBase.GetCurrentMethod()?.Name, name);
        return Properties.SingleOrDefault(c => c.Name == name);
    }

    internal object? GetValue(PropertyInfo prop, object obj)
    {
        try
        {
            Excel.Info("{$a}:{b} of {c}", this, MethodBase.GetCurrentMethod()?.Name, prop.Name);
            var value = prop.GetValue(obj);
            Excel.Debug("{$a}:{b} of {c} Value={d}", this, MethodBase.GetCurrentMethod()?.Name, prop.Name, value);
            return value;
        }
        catch (Exception ex)
        {
            Excel.Error(ex);
            return null;
        }
    }
    internal object? GetValue(PropertyInfo prop) => GetValue(prop, Item);
    public object? GetValue() => GetValue(Key)!;

    public object? GetValue(string name)
    {
        var prop = Property(name);
        if (prop is null)
        {
            Excel.Warning("{$a}: Property {b} does not exist!", this, name);
            return null;
        }
        return GetValue(prop);
    }

    internal Type GetObjectType()
    {
        Excel.Info("{$a} GetObjectType: {b}", this, Item.GetType().Name);
        return Item.GetType();
    }


}