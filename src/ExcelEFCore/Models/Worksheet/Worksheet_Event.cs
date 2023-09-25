namespace ExcelEFCore;

public partial class Worksheet
{
    internal event Action<IEnumerable<Element>, PropertyInfo> AddEvent;
    //Event parameters:
    //What are the affected elements?
    //What DbSet they belong?

    internal event Action<IEnumerable<Element>, PropertyInfo, PropertyInfo> RemoveEvent;
    //Event parameters:
    //What are the affected elements?
    //What DbSet they belong?
    //What is the Index property?

    internal event Action<IEnumerable<Element>, PropertyInfo, PropertyInfo> UpdateEvent;
    //Event parameters:
    //What are the affected elements?
    //What DbSet they belong?
    //What is the Index property?

    internal event Action<PropertyInfo> ClearAllEvent;
    //Event parameters:
    //What DbSet they belong?


}