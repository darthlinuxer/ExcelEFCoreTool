namespace ExcelEFCore;

internal partial class ContextHandler
{
    private readonly DbContext context;

    internal ContextHandler(DbContext context)
    {
        this.context = context;
    }

    internal bool AddElements(IEnumerable<Element> wkSheetElements, PropertyInfo dbSetProp)
    {
        try
        {
            var elementsToAdd = wkSheetElements.Select(c => c.Item);
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetProp.Name);
            contextElements.ToList().AddRange(elementsToAdd);
            elementsToAdd.ToList().ForEach(c => context.Entry(c).State = EntityState.Added);
            var result = context.SaveChanges() > 0;
            Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return false;
        }
    }

    internal bool RemoveElements(IEnumerable<Element> wkSheetElements, PropertyInfo dbSetProp, PropertyInfo key)
    {
        try
        {
            var elementsToRemove = wkSheetElements.Select(c => c.Item);
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = dbSetProp.GetGetMethod()?.Invoke(context, null) as IEnumerable<object>;
            contextElements = GetElements(dbSetProp.Name);
            contextElements?.ToList().RemoveAll(c =>
                      {
                          var contextId = GetValue(key, c);
                          var matchedElement = elementsToRemove.SingleOrDefault(d => GetValue(key, d)?.Equals(contextId ?? -1) ?? false);
                          if (matchedElement is not null) context.Entry(c).State = EntityState.Deleted;
                          return true;
                      });
            var result = context.SaveChanges() > 0;
            Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return false;
        }
    }

    internal bool UpdateElements(IEnumerable<Element> wkSheetElements, PropertyInfo dbSetProp, PropertyInfo key)
    {
        try
        {
            var elementsToUpdate = wkSheetElements.Select(c => c.Item);
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetProp.Name);
            contextElements?.ToList().ForEach(c =>
                       {
                           var indexValue = GetValue(key, c);
                           var updatedElement = elementsToUpdate.SingleOrDefault(d => GetValue(key, d)?.Equals(indexValue ?? -1) ?? false);
                           if (updatedElement is not null)
                           {
                               Copy(updatedElement, ref c);
                               context.Entry(c).State = EntityState.Modified;
                           }
                       });
            var result = context.SaveChanges() > 0;
            Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    internal bool ClearAll(PropertyInfo dbSetProp)
    {
        try
        {
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetProp.Name) as List<object>;
            contextElements?.Clear();
            var result = context.SaveChanges() > 0;
            Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return false;
        }

    }

    internal void Save() => context.SaveChanges();

    internal void Dispose() => context.Dispose();

}