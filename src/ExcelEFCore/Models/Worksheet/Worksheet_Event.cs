namespace ExcelEFCore;

public partial class Worksheet
{
    internal event Func<IEnumerable<Element>, PropertyInfo, bool> AddEvent;
    //Event parameters:
    //What are the affected elements?
    //What DbSet they belong?

    internal event Func<IEnumerable<Element>, PropertyInfo, PropertyInfo, bool> RemoveEvent;
    //Event parameters:
    //What are the affected elements?
    //What DbSet they belong?
    //What is the Index property?

     internal event Func<IEnumerable<Element>, PropertyInfo, PropertyInfo, bool> UpdateEvent;
    //Event parameters:
    //What are the affected elements?
    //What DbSet they belong?
    //What is the Index property?

     internal event Func<PropertyInfo, bool> ClearAllEvent;
    //Event parameters:
    //What DbSet they belong?


}