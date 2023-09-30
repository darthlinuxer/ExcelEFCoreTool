using ExcelEFCore;
using Model;

namespace Cli;

public class Program
{
    public static void Main(string[] args)
    {
        IExcelDbContext context = new AppDbContext("inmemoryDb.db");
        var rootCommand = new RootCommand();
        var createCommand = new Command("create", "Creates an Excel file");
        var fileArg = new Argument<string>("file", "Name of the Excel File");
        var logLevelOption = new Option<string>("--loglevel","LogLevel: VRB, DBG, INF, WRN, ERR, FTL");

        createCommand.AddArgument(fileArg);
        createCommand.AddOption(logLevelOption);

        createCommand.SetHandler((file, LogLevel) =>
        {
            var excel = Excel.Create(file: file, dbContext: context, LogLevel);
            excel!.ExportContextToWorksheets();
        }, fileArg, logLevelOption);

        var updateCommand = new Command("update",
                @"Updates Db according to Color Codes in Excel Table
                  1. Yellow Background (RGB 255,255,0): Update data
                  2. Blue Background (RGB 0,0,255): Adds new data
                  3. Red Background (RGB 255,0,0): Deletes Data");

        updateCommand.SetHandler((file) =>
        {
            var excel = Excel.Create(file: file, dbContext: context);
            excel!.ProcessColoredWorksheetToContext();
        }, fileArg);

        updateCommand.AddArgument(fileArg);

        var rebuildCommand = new Command("rebuild", "Refresh a worksheet with data from Db");
        var folderOpt = new Option<string>(new string[] { "-ws", "--worksheet" }, "Worksheet tab name");
        rebuildCommand.SetHandler((file, folder) =>
        {
            var excel = Excel.Create(file: file, dbContext: context);
            excel!.ExportContextToWorksheet(folder);
        }, fileArg, folderOpt);

        rebuildCommand.AddArgument(fileArg);
        rebuildCommand.AddOption(folderOpt);

        rootCommand.AddCommand(createCommand);
        rootCommand.AddCommand(updateCommand);
        rootCommand.AddCommand(rebuildCommand);
        rootCommand.Invoke(args);
    }

}


