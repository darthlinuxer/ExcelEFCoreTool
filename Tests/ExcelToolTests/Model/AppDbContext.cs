namespace Test;
public class AppDbContext : DbContext, IDisposable
{
    private string dbName = "";
    public DbSet<Person>? Persons { get; set; }
    public DbSet<Book>? Books { get; set; }
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
    public AppDbContext() { }
    public AppDbContext(string inMemoryDbName) => dbName = inMemoryDbName;
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder) => optionsBuilder.UseInMemoryDatabase(dbName);

    public override void Dispose() { Console.WriteLine("************** Context being disposed ************"); } // base.Dispose(); GC.SuppressFinalize(this); } 
}