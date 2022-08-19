namespace Test;

[TestClass]
public partial class ExcelEFCoreToolTests
{
    Excel? excel = null;
    AppDbContext? context = null;
    static int test = 0;

    public ExcelEFCoreToolTests()
    {

    }

    [TestInitialize]
    public void SetupTest()
    {
        //In Memory Entity Framework context leaks between tests unless it has a different db Name;
        context = new AppDbContext($"dbName{test}");
        excel = Excel.Create($"test{test}.xlsx", context, "INF");

        //Arrange
        var persons = new List<Person>
        {
            new Person(){Name = "Anakin SkyWalker"},
            new Person(){Name = "Luke Skywalker"}
        };

        var books = new List<Book>
        {
            new Book(){Name = "1001 Nights",PersonFK=1}
        };

        test++;

        context.Persons!.AddRange(persons);
        context.SaveChanges();
        context.Books!.AddRange(books);
        context.SaveChanges();
    }

    [TestCleanup]
    public void CleanTest()
    {
        excel!.Dispose();
    }



}