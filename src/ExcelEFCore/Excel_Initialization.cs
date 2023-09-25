
namespace ExcelEFCore;

public partial class Excel : IDisposable
{
    private ExcelPackage? app;
    private ExcelWorkbook? workBook;
    private ContextHandler? _contextHandler;
    private const string toolVersion = "1.0.0";
    private LoggerConfiguration loggerConfiguration = new LoggerConfiguration();
    private static Logger? log;
    private LoggingLevelSwitch levelSwitch = new LoggingLevelSwitch();
    public List<Worksheet> Worksheets { get; set; } = new List<Worksheet>();

    public ContextHandler? ContextHandler { get => _contextHandler; private set { _contextHandler = value; } }


    private Excel(string file, ContextHandler contextHandler, string minimumLevel = "Debug")
    {
        LogEventLevel eventLevel = LogEventLevel.Debug;
        if (minimumLevel == "Verbose") eventLevel = LogEventLevel.Verbose;
        if (minimumLevel == "VRB") eventLevel = LogEventLevel.Verbose;
        if (minimumLevel == "Debug") eventLevel = LogEventLevel.Debug;
        if (minimumLevel == "DBG") eventLevel = LogEventLevel.Debug;
        if (minimumLevel == "Information") eventLevel = LogEventLevel.Information;
        if (minimumLevel == "INF") eventLevel = LogEventLevel.Information;
        if (minimumLevel == "Warning") eventLevel = LogEventLevel.Warning;
        if (minimumLevel == "WRG") eventLevel = LogEventLevel.Warning;
        if (minimumLevel == "Error") eventLevel = LogEventLevel.Error;
        if (minimumLevel == "ERR") eventLevel = LogEventLevel.Error;
        if (minimumLevel == "Fatal") eventLevel = LogEventLevel.Fatal;
        if (minimumLevel == "FTL") eventLevel = LogEventLevel.Fatal;

        levelSwitch.MinimumLevel = eventLevel;
        log = loggerConfiguration
              .Enrich.WithProperty("Version", toolVersion)
              .Enrich.FromLogContext()
              .MinimumLevel.ControlledBy(levelSwitch)
              .WriteTo.Console()
              .WriteTo.File("log.txt", rollingInterval: RollingInterval.Hour)
              .CreateLogger();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ContextHandler = contextHandler;

        try
        {
            this.app = new ExcelPackage(file);
            this.workBook = app.Workbook;
            InitAligments();
            Excel.Info("{a}: Initialized ExcelToolEFIntegration version {toolVersion}", MethodBase.GetCurrentMethod()?.Name, toolVersion);
            Save();
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{a} {b} {c} ", MethodBase.GetCurrentMethod()?.Name, ex.Message, ex.InnerException?.Message);
        }
    }

    ~Excel()
    {
        Dispose();
    }

    public void Save()
    {
        try
        {
            _contextHandler?.Save();
            app?.Save();
            Excel.Info("{$a} {b}", this, MethodBase.GetCurrentMethod()?.Name);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{$a} {b} {c}", this, MethodBase.GetCurrentMethod()?.Name, ex.Message);
        }
    }
    public void Dispose()
    {
        Worksheets.ForEach(c =>
        {
            if (_contextHandler is not null)
            {
                c.AddEvent -= _contextHandler.AddElements;
                c.RemoveEvent -= _contextHandler.RemoveElements;
                c.UpdateEvent -= _contextHandler.UpdateElements;
                c.ClearAllEvent -= _contextHandler.ClearAll;
            }
        });
        this._contextHandler?.Dispose();
        this.app?.Dispose();
        GC.Collect();
    }

    public static Excel? Create(string? file, IDbContext dbContext, string eventLevel = "Debug", int indexCellColNumber = 1)
    {
        try
        {
            var contextHandler = new ContextHandler(dbContext);
            Excel.Info("Excel Static: {a} {$context} {$eventLevel} {indexCellColNumber}", MethodBase.GetCurrentMethod()?.Name, dbContext, eventLevel, indexCellColNumber);
            if (file is null || file == "")
            {
                Excel.Warning("File name: Default.xlsx");
                file = "Default.xlsx";
            }
            return new Excel(file, contextHandler, eventLevel);
        }
        catch (Exception ex)
        {
            Excel.Error(ex, "{b} {c}", MethodBase.GetCurrentMethod()?.Name, ex.Message);
            return null;
        }
    }

    public static void Debug(string msg, params object?[]? p)
    {
        log?.Debug(msg, p);
    }

    public static void Info(string msg, params object?[]? p)
    {
        log?.Information(msg, p);
    }

    public static void Warning(string msg, params object?[]? p)
    {
        log?.Warning(msg, p);
    }

    public static void Error(Exception ex, string? msgTemplate = null, params object?[]? args)
    {
        if (msgTemplate is null) msgTemplate = "{$a}";
        log?.Error(ex, msgTemplate, args);
    }

    public void Log(Exception ex)
    {
        log?.Error(ex, "{a}", ex.Message);
    }

}