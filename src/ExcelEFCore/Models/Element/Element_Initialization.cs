namespace ExcelEFCore;

public partial class Element
{
    internal object Item { get; init; }
    internal IEnumerable<PropertyInfo> Properties { get; init; }
    public PropertyInfo Key { get; init; }
    public CultureInfo Culture { get; init; }
    private Element(object item, string key, CultureInfo? culture = null)
    {
        try
        {
            Excel.Debug("{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, item);
            Culture = culture ?? new CultureInfo("en-US");
            Item = item;
            Properties = GetProperties(item);
            var keyProp = Property(key);
            if (keyProp is null) throw new Exception("{key} is not a valid property!");
            Key = keyProp;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    public void Deconstruct(out string keyName,
                           out object? keyvalue)
    {
        keyName = this.Key.Name;
        keyvalue = this.Key.GetValue(Item);
    }

    public static Element Factory(object item, string keyPropertyName, CultureInfo? culture = null)
    {
        return new Element(item, keyPropertyName, culture);
    }

    public override bool Equals(object? obj)
    {
        try
        {
            if (obj is null) return false;
            return obj is Element element &&
                   Properties.Select(c => new { value = GetValue(c) }).SequenceEqual(GetProperties(obj).Select(c => new { value = GetValue(c) }));
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}: {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return false;
        }
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Item, Properties, Key, Culture);
    }

    public Dictionary<string, object?> UnMatchedProperties(Element? element, bool compareId = false)
    {
        var result = new Dictionary<string, object?>();
        foreach (var prop in Properties)
        {
            if (prop.Name == this.Key.Name && compareId == false) continue;
            var sourceValue = prop.GetValue(Item);
            if (element is null) { result.Add(prop.Name, null); continue; }
            var targetValue = prop.GetValue(element.Item);
            if (sourceValue is not null && targetValue is null) { result.Add(prop.Name, null); continue; }
            if (sourceValue is not null && targetValue is not null) if (!(sourceValue.Equals(targetValue))) { result.Add(prop.Name, targetValue); continue; }
            if (sourceValue is null && targetValue is not null) { result.Add(prop.Name, targetValue);  }
        }
        return result;
    }

    public static bool operator ==(Element a, Element? b) => a.Equals(b);
    public static bool operator !=(Element a, Element? b) => a.Equals(b) is false;


}