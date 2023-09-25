namespace ExcelEFCore;

public interface IDbContext
{
    void EntryStatus(object entity, EntityState state);

    int SaveChanges();

    void Dispose();

}