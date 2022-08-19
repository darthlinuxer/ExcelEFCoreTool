namespace ExcelEFCore;

public partial class Excel
{
    private void InitAligments()
    {
        try
        {

            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var excelWorksheets = this.workBook?.Worksheets;
            foreach (var dbSet in contextHandler.GetAllDbSets())
            {
                var existingExcelWorksheet = excelWorksheets?.ToList().FirstOrDefault(c => c.Name == dbSet.Name);
                if (existingExcelWorksheet is not null)
                {
                    var existingWorksheet = Worksheet.Factory(dbSet, existingExcelWorksheet, true, "Id");
                    existingWorksheet.AddEvent += this.contextHandler.AddElements;
                    existingWorksheet.RemoveEvent += this.contextHandler.RemoveElements;
                    existingWorksheet.UpdateEvent += this.contextHandler.UpdateElements;
                    existingWorksheet.ClearAllEvent += this.contextHandler.ClearAll;
                    Worksheets.Add(existingWorksheet);
                    Save();
                    continue;
                }
                var newExcelWorksheet = this.workBook?.Worksheets.Add(dbSet.Name)!;
                Save();
                var newWorksheet = Worksheet.Factory(dbSet, newExcelWorksheet, true, "Id");
                newWorksheet.AddEvent += this.contextHandler.AddElements;
                newWorksheet.RemoveEvent += this.contextHandler.RemoveElements;
                newWorksheet.UpdateEvent += this.contextHandler.UpdateElements;
                newWorksheet.ClearAllEvent += this.contextHandler.ClearAll;
                Worksheets.Add(newWorksheet);
            }
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    public void ExportContextToWorksheets(CultureInfo? culture = null)
    {
        try
        {
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            foreach (var worksheet in Worksheets)
            {
                Excel.Info("{$a} {b} Exporting context to worksheet {c}", this, MethodBase.GetCurrentMethod()?.Name, worksheet.Name);
                worksheet.ImportFromContext(this.contextHandler, culture);
            }
            Save();
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    public void ExportContextToWorksheet(string name, CultureInfo? culture = null)
    {
        try
        {
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            var worksheet = Worksheets.FirstOrDefault(c => c.Name == name);
            if (worksheet is null) return;
            Excel.Info("{$a} {b} Exporting context to worksheet {c}", this, MethodBase.GetCurrentMethod()?.Name, worksheet.Name);
            worksheet.ImportFromContext(this.contextHandler, culture);
            Save();
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }

    public void ProcessColoredWorksheetToContext(CultureInfo? culture = null)
    {
        try
        {
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
            foreach (var worksheet in Worksheets)
            {
                Excel.Info("{$a} {b} Importing worksheet to context {c}", this, MethodBase.GetCurrentMethod()?.Name, worksheet.Name);
                worksheet.ProcessColoredToContext();
            }
            Save();

        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }
}