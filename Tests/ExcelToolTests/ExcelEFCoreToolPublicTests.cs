namespace Test;


public partial class ExcelEFCoreToolTests
{


    [TestMethod]
    public void ExcelWithDb_ExportContextToWorksheets_MustCreateWorksheetsWithData()
    {
        //Arrange

        //Act
        excelWithDb!.ExportContextToWorksheets(new System.Globalization.CultureInfo("en-US"));
        var persons = excelWithDb.GetElementsFromWorksheet<Person>("Persons");
        var books = excelWithDb.GetElementsFromWorksheet<Book>("Books");

        //Assert
        Assert.IsTrue(persons?.Count() == 2);
        Assert.IsTrue(books?.Count() == 1);
    }

    [TestMethod]
    public void ExcelWithDb_UpdateWorksheetToContext_MustUpdateContext()
    {
        //Arrange
        excelWithDb!.ExportContextToWorksheets(new System.Globalization.CultureInfo("en-US"));
        excelWithDb!.Save();

        var vader = new Person() { Id = 1, Name = "Darth Vader" }; //currently on context as Anakin
        var updatedElement = Element.Factory(vader, "Id");
        var worksheet = excelWithDb!.Worksheets.FirstOrDefault(c => c.Name == "Persons");
        worksheet?.Compare(e => e.GetValue()!.Equals(1), updatedElement, Color.Yellow, true);
        worksheet?.UpdateWorksheetOnly(updatedElement); // Now anakin changed to Vader and the name is yellow 
        excelWithDb!.Save();

        var camilo = new Person() { Name = "Camilo" };
        var newElement = Element.Factory(camilo, "Id");
        worksheet?.Add(newElement);
        worksheet?.Compare(e => e.GetValue("Name")!.Equals("Camilo"), null, Color.Blue, true);
        worksheet?.Compare(e => e.GetValue("Name")!.Equals("Luke Skywalker"), null, Color.Red, true);
        excelWithDb!.Save();

        //Act
        excelWithDb!.ProcessColoredWorksheetToContext(new System.Globalization.CultureInfo("en-US"));
        excelWithDb!.Save();
        excelWithDb!.ExportContextToWorksheets();
        var elements = excelWithDb!.GetElementsFromWorksheet<Person>("Persons");

        //Assert
        Assert.IsTrue(elements?.FirstOrDefault(c => c.Name == "Luke Skywalker") is null);
        Assert.IsTrue(elements?.FirstOrDefault(c => c.Name == vader.Name) is not null);
        Assert.IsTrue(elements?.FirstOrDefault(c => c.Name == camilo.Name) is not null);
        Assert.IsTrue(elements?.Count() == 2);

    }

    [TestMethod]
    public void ExcelWithApi_ExportContextToWorksheets_MustCreateWorksheetsWithData()
    {
        //Arrange

        //Act
        excelWithApi!.ExportContextToWorksheets(new System.Globalization.CultureInfo("en-US"));
        var persons = excelWithApi.GetElementsFromWorksheet<Person>("Persons");
        var books = excelWithApi.GetElementsFromWorksheet<Book>("Books");

        //Assert
        Assert.IsTrue(persons?.Count() == 2);
        Assert.IsTrue(books?.Count() == 1);
    }

    [TestMethod]
    public void ExcelWithApi_UpdateWorksheetToContext_MustUpdateContext()
    {
        //Arrange
        excelWithApi!.ExportContextToWorksheets(new System.Globalization.CultureInfo("en-US"));
        excelWithApi!.Save();

        var vader = new Person() { Id = 1, Name = "Darth Vader" }; //currently on context as Anakin
        var updatedElement = Element.Factory(vader, "Id");
        var worksheet = excelWithApi!.Worksheets.FirstOrDefault(c => c.Name == "Persons");
        worksheet?.Compare(e => e.GetValue()!.Equals(1), updatedElement, Color.Yellow, true);
        worksheet?.UpdateWorksheetOnly(updatedElement); // Now anakin changed to Vader and the name is yellow 
        excelWithApi!.Save();

        var camilo = new Person() { Name = "Camilo" };
        var newElement = Element.Factory(camilo, "Id");
        worksheet?.Add(newElement);
        worksheet?.Compare(e => e.GetValue("Name")!.Equals("Camilo"), null, Color.Blue, true);
        worksheet?.Compare(e => e.GetValue("Name")!.Equals("Luke Skywalker"), null, Color.Red, true);
        excelWithApi!.Save();

        //Act
        excelWithApi!.ProcessColoredWorksheetToContext(new System.Globalization.CultureInfo("en-US"));
        excelWithApi!.Save();
        excelWithApi!.ExportContextToWorksheets();
        var elements = excelWithApi!.GetElementsFromWorksheet<Person>("Persons");

        //Assert
        Assert.IsTrue(elements?.FirstOrDefault(c => c.Name == "Luke Skywalker") is null);
        Assert.IsTrue(elements?.FirstOrDefault(c => c.Name == vader.Name) is not null);
        Assert.IsTrue(elements?.FirstOrDefault(c => c.Name == camilo.Name) is not null);
        Assert.IsTrue(elements?.Count() == 2);

    }

}