namespace ExcelEFCore;

public partial class ContextHandler
{
    private readonly IDbContext _context;

    public ContextHandler(IDbContext context)
    {
        this._context = context;
    }

    public void AddElements(IEnumerable<Element> wkSheetElements, PropertyInfo dbSetProp)
    {
        try
        {
            var elementsToAdd = wkSheetElements.Select(c => c.Item);
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetProp.Name);
            contextElements.ToList().AddRange(elementsToAdd);
            elementsToAdd.ToList().ForEach(c => _context.EntryStatus(c, EntityState.Added));

            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            //return false;
        }
    }

    public void AddElements<T>(IEnumerable<T> collection, string dbSetName) where T : class
    {
        try
        {
            var elementsToAdd = collection;
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetName);
            contextElements.ToList().AddRange(elementsToAdd);
            elementsToAdd.ToList().ForEach(c => _context.EntryStatus(c, EntityState.Added));
            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            // return false;
        }
    }

    public void RemoveElements(IEnumerable<Element> wkSheetElements, PropertyInfo dbSetProp, PropertyInfo key)
    {
        try
        {
            var elementsToRemove = wkSheetElements.Select(c => c.Item);
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetProp.Name);
            contextElements?.ToList().RemoveAll(c =>
                      {
                          var contextId = GetValue(key, c);
                          var matchedElement = elementsToRemove.SingleOrDefault(d => GetValue(key, d)?.Equals(contextId ?? -1) ?? false);
                          if (matchedElement is not null) _context.EntryStatus(c, EntityState.Deleted);
                          return true;
                      });
            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            // return false;
        }
    }

    public void RemoveElements<T>(IEnumerable<T> collection, string dbSetName, string key) where T : class
    {
        try
        {
            var elementsToRemove = collection;
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetName);
            var keyProp = GetProperty(collection.First(), key);
            if (keyProp is null) throw new Exception("Key is not a Property of the item in the collection!");
            contextElements?.ToList().RemoveAll(c =>
                      {
                          var contextId = GetValue(keyProp!, c);
                          var matchedElement = elementsToRemove.SingleOrDefault(d => GetValue(keyProp!, d)?.Equals(contextId ?? -1) ?? false);
                          if (matchedElement is not null) _context.EntryStatus(c, EntityState.Deleted);
                          return true;
                      });
            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            // return false;
        }
    }


    public void UpdateElements(IEnumerable<Element> wkSheetElements, PropertyInfo dbSetProp, PropertyInfo key)
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
                               _context.EntryStatus(c, EntityState.Modified);
                           }
                       });
            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    public void UpdateElements<T>(IEnumerable<T> collection, string dbSetName, string key) where T : class
    {
        try
        {
            var elementsToUpdate = collection;
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetName);
            var keyProp = GetProperty(collection.First(), key);
            if (keyProp is null) throw new Exception("Key is not a Property of the item in the collection!");
            contextElements?.ToList().ForEach(c =>
                       {
                           var indexValue = GetValue(keyProp, c);
                           var updatedElement = elementsToUpdate.SingleOrDefault(d => GetValue(keyProp, d)?.Equals(indexValue ?? -1) ?? false);
                           if (updatedElement is not null)
                           {
                               Copy(updatedElement, ref c);
                               _context.EntryStatus(c, EntityState.Modified);
                           }
                       });
            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    public void ClearAll(PropertyInfo dbSetProp)
    {
        try
        {
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetProp.Name) as List<object>;
            contextElements?.Clear();
            contextElements?.ToList().RemoveAll(c =>
            {
                _context.EntryStatus(c, EntityState.Deleted);
                return true;
            });
            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            //return false;
        }
    }

    public void ClearAll(string dbSetName)
    {
        try
        {
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var contextElements = GetElements(dbSetName) as List<object>;
            contextElements?.Clear();
            contextElements?.ToList().RemoveAll(c =>
            {
                _context.EntryStatus(c, EntityState.Deleted);
                return true;
            });
            // var result = _context.SaveChanges() > 0;
            // Excel.Debug("{$a} {b} result={c}", this, MethodBase.GetCurrentMethod()?.Name, result);
            // return result;
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a}:{b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            //return false;
        }
    }

    public void Save() => _context.SaveChanges();

    public void Dispose() => _context.Dispose();

}