namespace Test;


public partial class ExcelEFCoreToolTests
{


    [TestMethod]
    public void Excel_ExportContextToWorksheets_MustCreateWorksheetsWithData()
    {
        //Arrange

        //Act
        excel!.ExportContextToWorksheets(new System.Globalization.CultureInfo("en-US"));
        var persons = excel.GetElementsFromWorksheet<Person>("Persons");
        var books = excel.GetElementsFromWorksheet<Book>("Books");

        //Assert
        Assert.IsTrue(persons?.Count() == 2);
        Assert.IsTrue(books?.Count() == 1);
    }

    [TestMethod]
    public void Excel_UpdateWorksheetToContext_MustUpdateContext()
    {
        //Arrange
        excel!.ExportContextToWorksheets(new System.Globalization.CultureInfo("en-US"));
        excel!.Save();

        var vader = new Person() { Id = 1, Name = "Darth Vader" }; //currently on context as Anakin
        var updatedElement = Element.Factory(vader, "Id");
        var worksheet = excel!.Worksheets.FirstOrDefault(c => c.Name == "Persons");
        worksheet?.Compare(e => e.GetValue()!.Equals(1), updatedElement, Color.Yellow, true);
        worksheet?.UpdateWorksheetOnly(updatedElement); // Now anakin changed to Vader and the name is yellow 
        excel!.Save();

        var camilo = new Person() { Name = "Camilo" };
        var newElement = Element.Factory(camilo, "Id");
        worksheet?.Add(newElement);
        worksheet?.Compare(e=>e.GetValue("Name")!.Equals("Camilo"), null, Color.Blue, true);
        worksheet?.Compare(e=>e.GetValue("Name")!.Equals("Luke Skywalker"), null, Color.Red, true);
        excel!.Save();

        //Act
        excel!.ProcessColoredWorksheetToContext(new System.Globalization.CultureInfo("en-US"));
        excel!.Save();
        excel!.ExportContextToWorksheets();
        var elements = excel!.GetElementsFromWorksheet<Person>("Persons");

        //Assert
        Assert.IsTrue(elements?.FirstOrDefault(c=>c.Name=="Luke Skywalker") is null);
        Assert.IsTrue(elements?.FirstOrDefault(c=>c.Name==vader.Name) is not null);
        Assert.IsTrue(elements?.FirstOrDefault(c=>c.Name==camilo.Name) is not null);
        Assert.IsTrue(elements?.Count()==2);

    }

}