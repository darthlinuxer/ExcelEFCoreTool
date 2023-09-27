using ExcelEFCore;
using Microsoft.EntityFrameworkCore;

namespace Model;
public class AppDbContext : DbContext, IDisposable, IExcelDbContext
{
    private string dbName = "";
    public DbSet<Person>? Persons { get; set; }
    public DbSet<Book>? Books { get; set; }
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
    public AppDbContext() { }
    public AppDbContext(string inMemoryDbName) => dbName = inMemoryDbName;
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder) => optionsBuilder.UseInMemoryDatabase(dbName);

    public override void Dispose()
    {
        Console.WriteLine("************** Context being disposed ************");
    }

    public void EntryStatus(object entity, EntityState state)
    {
        Entry(entity).State = state;
    }
}