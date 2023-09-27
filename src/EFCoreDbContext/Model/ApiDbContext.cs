using ExcelEFCore;
using Microsoft.EntityFrameworkCore;

namespace Model;

//Imagine this is not an InMemory List but an API to store collection in Database
//Why it is needed, because EFCore DbContexts needs a Direct Conection with a Database via Connection String
//And with this implementation, you can communicate with the Server API and give this class to the ContextHandler
//to update the Excel Tables
public class ApiDbContext : IExcelDbContext
{
    private List<(Person, EntityState)> PersonOperations { get; set; } = new List<(Person, EntityState)>();
    private List<(Book, EntityState)> BookOperations { get; set; } = new List<(Book, EntityState)>();
    public List<Person>? Persons { get; set; } = new List<Person>();
    public List<Book>? Books { get; set; } = new List<Book>();
    public void Dispose()
    {
        Persons = null;
        Books = null;
    }

    public void EntryStatus(object entity, EntityState state)
    {
        if (entity is Person) PersonOperations.Add(((Person)entity, state));
        if (entity is Book) BookOperations.Add(((Book)entity, state));
    }


    public int SaveChanges()
    {
        var count = 0;
        foreach (var personOp in PersonOperations)
        {
            if (personOp.Item2 == EntityState.Added) Persons!.Add(personOp.Item1);
            if (personOp.Item2 == EntityState.Modified && Persons!.Exists(c => c?.Id == personOp.Item1.Id))
            {
                Persons.RemoveAll(c => c.Id == personOp.Item1.Id);
                Persons.Add(personOp.Item1);
            }
            if (personOp.Item2 == EntityState.Deleted) Persons!.Remove(personOp.Item1);
            count++;
        }
        foreach (var booksOp in BookOperations)
        {
            if (booksOp.Item2 == EntityState.Added) Books!.Add(booksOp.Item1);
            if (booksOp.Item2 == EntityState.Modified && Books!.Exists(c => c?.Id == booksOp.Item1.Id))
            {
                Books.RemoveAll(c => c.Id == booksOp.Item1.Id);
                Books.Add(booksOp.Item1);
            }
            if (booksOp.Item2 == EntityState.Deleted) Books!.Remove(booksOp.Item1);
            count++;
        }
        PersonOperations.Clear();
        BookOperations.Clear();
        return count;
    }
}