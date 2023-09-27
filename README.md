# ExcelEFCoreTool
  Integration between Excel and Entity Framework Core  
  With the ExcelEFCore DLL Library you can:
  * Automatically create an Excel file containing all data inside the DbSets of an EFCore DbContext
  * Bulk Updates on the Excel file based on background colors of the rows in the Worksheet
    * Yellow RGB (255,255,0): Update the Record
    * Red RGB (255,0,0): Delete the Record
    * Blue RGB (0,0,255): Add the Record

Note: Remember to save the Excel file and close it after applying the colors.


## Projects inside this Repo
1. EFCoreDbContext: A simple EFCore DbContext with 2 DbSets: Person and Books
2. ExcelEFCore: The Library .DLL that interfaces with Excel
3. ExcelEFCore Cli: A Cli to run commands on the Library

## Tests
Unit Tests will reference the DbContext and will create a temporary Excel file to run automated CRUDS. 
Check the unit tests to see how to use the library in your projects
```
  dotnet test
```

## CLI
The cli project has a template code that you can use to create your own cli to automatically bulk update Excel files