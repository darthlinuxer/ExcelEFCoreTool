using System.Dynamic;

namespace ExcelEFCore;

internal partial class ContextHandler
{
    internal object? GetValue(PropertyInfo prop, object obj)
    {
        try
        {
            Excel.Info("{$a}:{b} of {c}", this, MethodBase.GetCurrentMethod()?.Name, prop.Name);
            var value = prop.GetValue(obj);
            Excel.Debug("{$a}:{b} Value={c}", this, MethodBase.GetCurrentMethod()?.Name, value);
            return value;
        }
        catch (Exception ex)
        {
            Excel.Error(ex);
            return null;
        }
    }

    internal PropertyInfo? GetDbSetProperty(ExcelWorksheet wksheet) => context.GetType().GetProperty(wksheet.Name, BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.Instance);
    internal PropertyInfo? GetDbSetProperty(string name) => context.GetType().GetProperty(name, BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.Instance);
    internal IEnumerable<PropertyInfo> GetAllDbSets() => context.GetType().GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.Instance);
    internal IEnumerable<PropertyInfo> GetProperties(object source) => source.GetType().GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.Instance);
    internal Type? GetUnderlyingType(PropertyInfo? dbSet) => dbSet?.PropertyType.GenericTypeArguments.Single().UnderlyingSystemType;
    internal object? GetObject(IEnumerable<object> elements, PropertyInfo prop, object value) => elements.ToList().FirstOrDefault(c => prop.GetValue(c) == value);

    internal IEnumerable<object> GetElements(string dbSetName)
    {
        try
        {
            Excel.Info("{a} {b} dbSet={c}", this, MethodBase.GetCurrentMethod()?.Name, dbSetName);
            var dbSet = GetDbSetProperty(dbSetName);
            if(dbSet is null) throw new Exception($"There is no dbSet called {dbSetName}");
            var elements = dbSet.GetGetMethod()!.Invoke(context, null) as IEnumerable<object>;
            Excel.Debug("{a} {b} dbSet={c} elements.Count()={d}", this, MethodBase.GetCurrentMethod()?.Name, dbSetName, elements?.Count());
            return elements!;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a} {b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    internal object Clone(object source)
    {
        var properties = GetProperties(source);
        var target = new ExpandoObject() as IDictionary<string, object?>;
        foreach (var property in properties)
        {
            target.Add(property.Name, property.GetValue(source));
        }
        return target;
    }

    internal void Copy(object source, ref object target)
    {
        var properties = GetProperties(source);
        foreach (var property in properties)
        {
            property.SetValue(target, property.GetValue(source));
        }
        
    }


}