using Model;

namespace Test;

[TestClass]
public partial class ExcelEFCoreToolTests
{
    Excel? excelWithDb = null;
    Excel? excelWithApi = null;
    static int test = 0;

    public ExcelEFCoreToolTests()
    {

    }

    [TestInitialize]
    public void SetupTest()
    {
        //In Memory Entity Framework context leaks between tests unless it has a different db Name;
        IExcelDbContext appDbContext = new AppDbContext($"dbName{test}");
        IExcelDbContext apiDbContext = new ApiDbContext();
        excelWithDb = Excel.Create($"test{test}.xlsx", appDbContext, "INF");
        excelWithApi = Excel.Create($"test{test}.xlsx", apiDbContext, "INF");

        //Arrange
        var persons = new List<Person>
        {
            new Person(){ Name = "Anakin Skywalker"},
            new Person(){ Name = "Luke Skywalker"}
        };

        var books = new List<Book>
        {
            new Book(){Name = "1001 Nighs", PersonFK = 1}
        };


        test++;

        var contexhandler = excelWithDb!.ContextHandler;
        contexhandler!.AddElements(persons, "Persons"!);
        contexhandler!.AddElements(books, "Books");
        contexhandler!.Save();

        var apicontexhandler = excelWithApi!.ContextHandler;
        apicontexhandler!.AddElements(persons, "Persons"!);
        apicontexhandler!.AddElements(books, "Books");
        apicontexhandler!.Save();
    }

    [TestCleanup]
    public void CleanTest()
    {
        excelWithDb!.Dispose();

    }



}