namespace ExcelEFCore;

public partial class Worksheet
{

    internal void ProcessColoredToContext()
    {
        Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
        var rowsToDelete = RowsToDelete();
        Excel.Debug("{$a}:{b} rowsToDelete.Count()={c}", this, MethodBase.GetCurrentMethod()?.Name, rowsToDelete?.Count());
        var elementsToDelete = new List<Element>();
        rowsToDelete?.ToList().ForEach(row =>
        {
            var element = GetElement(row);
            elementsToDelete.Add(element);
        });
        if (elementsToDelete.Count() > 0)
        {
            RemoveEvent?.Invoke(elementsToDelete, ContextDbSet, KeyProp);
            rowsToDelete?.ToList().ForEach(row => ColorRow(row, Color.Black));
        }

        var rowsToAdd = RowsToAdd();
        Excel.Debug("{$a}:{b} rowsToAdd.Count()={c}", this, MethodBase.GetCurrentMethod()?.Name, rowsToAdd?.Count());
        var elementsToAdd = new List<Element>();
        rowsToAdd?.ToList().ForEach(row =>
        {
            var element = GetElement(row);
            elementsToAdd.Add(element);
        });
        if (elementsToAdd.Count() > 0)
        {
            AddEvent?.Invoke(elementsToAdd, ContextDbSet);
            rowsToAdd?.ToList().ForEach(row =>
            {
                var indexofRow = rowsToAdd?.ToList().IndexOf(row);
                var element = elementsToAdd.ElementAt(indexofRow!.Value);
                WriteToRow(row, element);
                ColorRow(row, UISuccessColor);
            });
        }

        var rowsToUpdate = RowsToUpdate();
        var elementsToUpdate = new List<Element>();
        Excel.Debug("{$a}:{b} rowsToUpdate.Count()={c}", this, MethodBase.GetCurrentMethod()?.Name, rowsToUpdate?.Count());
        rowsToUpdate?.ToList().ForEach(row =>
        {
            var element = GetElement(row);
            elementsToUpdate.Add(element);
        });
        if (elementsToUpdate?.Count() > 0)
        {
            UpdateEvent?.Invoke(elementsToUpdate, ContextDbSet, KeyProp);
            rowsToUpdate?.ToList().ForEach(row =>
            {
                var indexofRow = rowsToUpdate?.ToList().IndexOf(row);
                var element = elementsToUpdate.ElementAt(indexofRow!.Value);
                WriteToRow(row, element);
                ColorRow(row, UISuccessColor);
            });
        }
    }

    internal void ExportToContext()
    {
        try
        {
            Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            var elements = new List<Element>();
            for (var row = 2; row <= GetLastRow(); row++)
            {
                var element = GetElement(row);
                elements.Add(element);
            }
            ClearAllEvent?.Invoke(ContextDbSet);
            AddEvent?.Invoke(elements, ContextDbSet);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

    internal void ImportFromContext(ContextHandler context, CultureInfo? culture = null)
    {
        try
        {
            Excel.Info("{$a}:{b}", this, MethodBase.GetCurrentMethod()?.Name);
            ClearAll(propagateClearToContext: false);
            var dbSet = context.GetDbSetProperty(RealWorksheet.Name);
            if (dbSet is null) throw new Exception($"Worksheet {RealWorksheet.Name} have no associated DbSet");
            IEnumerable<object>? contextElements = context.GetElements(dbSet!.Name);

            if (!ValidateHeaders()) { ClearAll(propagateClearToContext: false); CreateHeaders(); }

            if (contextElements is not null)
                foreach (var contextElement in contextElements)
                {
                    var element = Element.Factory(contextElement, KeyName, culture);
                    Excel.Debug("{$a}:{b} element={@c}", this, MethodBase.GetCurrentMethod()?.Name, element.Item);
                    Append(element);
                }

        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a}:{b} {@c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
            throw;
        }
    }

}