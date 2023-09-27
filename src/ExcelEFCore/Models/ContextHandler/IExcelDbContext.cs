namespace ExcelEFCore;

public interface IExcelDbContext
{
    void EntryStatus(object entity, EntityState state);

    int SaveChanges();

    void Dispose();

}