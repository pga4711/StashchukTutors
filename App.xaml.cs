using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.IO.Abstractions;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Serilog;
using SimpleInjector;
using VMI.VibLyze.AutoMapperConfig;
using VMI.VibLyze.Common.Utilities;
using VMI.VibLyze.Instruments;
using VMI.VibLyze.Resources;
using VMI.VibLyze.Services;
using VMI.VibLyze.Services.BearingLibraryRepository;
using VMI.VibLyze.Services.GraphCursorService;
using VMI.VibLyze.Services.Implementation;
using VMI.VibLyze.Services.Implementation.ClipboardService;
using VMI.VibLyze.Services.Implementation.Conversion;
using VMI.VibLyze.Services.Implementation.RouteFileService;
using VMI.VibLyze.Services.Implementation.Settings;
using VMI.VibLyze.Services.Implementation.ShaftSpeedService;
using VMI.VibLyze.Services.RouteFileService;
using VMI.VibLyze.Utilities;
using VMI.VibLyze.ViewModels.MainWindow;
using VMI.VibLyze.ViewModels.TemperatureTrendView;
using VMI.VibLyze.ViewModels.VibrationTrendView;
using VMI.VibLyze.Views;
using WupiEngine;

namespace VMI.VibLyze
{
    public partial class App : Application
    {
        public static ILogger Logger { get; private set; }
        private Container _container;

        [NotNull]
        private static string LocalStorageFolder { get; } = Path.Combine(
            Path.GetPathRoot(Environment.SystemDirectory)!, "VibLyze");

        [NotNull]
        private static string InstallFolder { get; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        static App()
        {
            const string lightningChartDeploymentKey = "lgCAAA+28YPCCtcBJABVcGRhdGVhYmxlVGlsbD0yMDIyLTAzLTE4I1JldmlzaW9uPTACgAYhL74O0Jr0m+CRqdGgY6pXbH9u+QFnEEFnB7bkk5m5q0XQ/yMhxYfqlCaa9/5s+u3IgHm2lCiSTpklPoJR9qiAUjaESCr/3TuL66kskrCcdqJfP8lHiKgYnEV+noCxFu6mH7J19LcmrfA43eFl+fxIKVnPWf3RkRNBFc7X4NPSKuYRqK2/UFbdWSyymhE0IVbmTP6AGCpsX71fdm8cDLPzuWZYZKLtssslj+4ymSOAdDfCBMS4aXbcZS/5zlDZua5l2OFkfuaAeqgTP6mUo3xLkPV/EK0ZJPzDCj/xbFnEXdJENHTaYlPVKotwfHCuZknGzDwb78ZcGxaDXsdK1XpCu8B7qQe9p2rIY3ajdxozHMXfCZcF7glfQa2XwQOuE7Sn9HkbH+aqg/oAVBCtCkm3kJ5cHcnqOul9eQmdEAfL1pJvqNdvBSzQ6FWTaFgg6o8ob4kcgnJX4DOHHFcmPV4F2gcLZ398/Gs/RFGU9yLGpxyYQ5HawCGn6IoY2fXhk4A=";
            Arction.Wpf.ChartingMVVM.LightningChart.SetDeploymentKey(lightningChartDeploymentKey);
        }

        public static void ShowLicenseErrorDialog()
        {
            UserMessage.UserMessage.ShowMessage(0, new Dictionary<string, object>() { { "Owner", App.Current.MainWindow } });
        }

        private static List<UserMessage.DialogViewModel.License> GetLicenses()
        {
            var licenses = new List<UserMessage.DialogViewModel.License>();
            var instruments = new List<IInstrument>() { new ViberX4(), new ViberX5() };

            foreach (var i in instruments)
            {
                var mask = Wupi.QueryInfo(i.Id, QueryInfoOption.BoxMask);
                var serial = Wupi.QueryInfo(i.Id, QueryInfoOption.BoxSerial);
                var productCode = Wupi.QueryInfo(i.Id, QueryInfoOption.ProductCode);
                var expirationTime = Wupi.QueryInfo(i.Id, QueryInfoOption.ExpirationTime);

                Logger.Debug($"Instrument {i.Name} License {i.Id} Mask {mask} Serial {serial} ProductCode {productCode} " +
                    $"ExpirationTime {expirationTime}");

                var l = new UserMessage.DialogViewModel.License()
                {
                    Id = i.Id,
                    Name = i.Name,
                    Expires = ""
                };
                if (expirationTime > 0)
                {
                    var et = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(expirationTime).ToLocalTime();
                    var el = $" ({et.ToShortDateString()})";
                    l.ExpiresDate = et;
                    l.Expires = el;
                }
                licenses.Add(l);
            }
            return licenses;
        }

        private static void ShowLicenseStatusDialog(List<UserMessage.DialogViewModel.License> licenses, string status = "")
        {
            UserMessage.UserMessage.ShowStatus(licenses, status);
        }

        private static void CheckLicenses()
        {
            var licenses = GetLicenses();
            if (UserMessage.UserMessage.LicensesAboutToExpire(licenses))
            {
                ShowLicenseStatusDialog(licenses, UM.Resources.Resources.LicenseExpiringWarnText);
            }
        }
        
        public static void ShowLicenseStatusDialog()
        {
            var licenses = GetLicenses();
            var status = UserMessage.UserMessage.LicensesAboutToExpire(licenses) ? UM.Resources.Resources.LicenseExpiringWarnText : "";
            UserMessage.UserMessage.ShowStatus(licenses, status);
        }

        private void OnResourceProviderDataChanged(object sender, EventArgs e)
        {
            if (Logger != null)
            {
                Logger.Debug("ResourceProviderDataChanged");
                // TODO: Trigger view update?
            }
        }

        private void OnLanguageSelected(object sender, RoutedEventArgs e)
        {
            var sl = e.OriginalSource as MenuItem;
            var c = CultureResources.SupportedCultures.Find(c => c.TwoLetterISOLanguageName == sl.Name);
            if (c != null)
            {
                Logger.Debug($"Selecting culture {c.Name} -- {c.TwoLetterISOLanguageName} -- {c.EnglishName} -- {c.NativeName}");
                if (CultureResources.ChangeCulture(c))
                {
                    var userSettings = _container.GetInstance<UserSettingsService>();
                    userSettings.Language = c.TwoLetterISOLanguageName;
                }
            }
        }

        private void AppStartup(object sender, StartupEventArgs e)
        {
            InitializeLogging();
            Logger = LogUtility.CreateClassLogger<App>();

            Logger.Information("Application startup - Version: {ProductVersion}", AssemblyInfo.ProductVersion);

            SetupUnhandledExceptionHandling();

            _container = CreateContainer();

            InitializeApplicationSettings();

            var userSettings = _container.GetInstance<UserSettingsService>();
            var c = CultureResources.SupportedCultures.Find(c => c.TwoLetterISOLanguageName == userSettings.Language);
            CultureInfo currentCulture = c ?? VibLyze.Properties.Settings.Default.DefaultCulture;
            CultureResources.ChangeCulture(c);
            CultureResources.ResourceProvider.DataChanged += new EventHandler(OnResourceProviderDataChanged);

            // Open last used (right now always the same) database
            OpenLastUsedDatabaseAsync().Forget();

            InitializeBearingLibraryAsync().Forget();

            ShowMainWindow(currentCulture);

            CheckLicenses();
        }

        private static void InitializeLogging()
        {
            Log.Logger = new LoggerConfiguration() 
                .MinimumLevel.Debug() // TODO: At least log-level should be configurable without recompile!
                .WriteTo.Async(sink => sink.File(
                    Path.Combine(LocalStorageFolder, "Logs", "VibLyze.log"),
                    rollOnFileSizeLimit: true,
                    fileSizeLimitBytes: 1024*1024*32,
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] [{SourceContext}] {Message:lj} {Properties}{NewLine}{Exception}"))
                .WriteTo.Debug(
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] [{SourceContext}] {Message:lj} {Properties}{NewLine}{Exception}")
                // Default template: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}"
                .CreateLogger();
        }

        private void SetupUnhandledExceptionHandling()
        {
            // Gotta catch' em all!

            // Unobserved exceptions on Tasks. 
            TaskScheduler.UnobservedTaskException += (_, args) =>
            {
                // This should not happen. When it does, there is a task that should be awaited, e.g. directly or using TaskExtensions.Forget.
                Logger.Error(args.Exception, "TaskScheduler.UnobservedTaskException");
                // We could do args.SetObserved(); here, but let it propagate and crash the application instead. I.e., fail fast.
            };

            // Unhandled exceptions on the UI thread
            DispatcherUnhandledException += (_, args) =>
            {
                Logger.Error(args.Exception, "Application.Current.DispatcherUnhandledException");
                MessageBox.Show($"{VibLyze.Resources.Resources.UnexpectedErrorMessage}\n\n" +
                                string.Format(VibLyze.Resources.Resources.DetailsMessage, args.Exception),
                    VibLyze.Resources.Resources.UnexpectedErrorDialogCaption, MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(-1);
                args.Handled = true;
            };

            // Last resort. If we get this, the application will crash.
            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                Logger.Fatal(args.ExceptionObject as Exception, "AppDomain.CurrentDomain.UnhandledException");
                MessageBox.Show($"{VibLyze.Resources.Resources.UnexpectedFatalErrorMessage}\n\n" +
                                string.Format(VibLyze.Resources.Resources.DetailsMessage, args.ExceptionObject as Exception),
                    VibLyze.Resources.Resources.UnexpectedErrorDialogCaption, MessageBoxButton.OK, MessageBoxImage.Error);
            };
        }

        private static Container CreateContainer()
        {
            var container = new Container();

            container.RegisterSingleton<IServiceProvider>(() => container);
            container.RegisterSingleton<IDatabaseService, DatabaseService>();
            container.RegisterSingleton<IModelChangeEventBroker, ModelChangedEventBroker>();
            container.Register<IJetDbConverterService, JetDbConverterService>();
            container.Register<IRouteFileService, RouteFileService>();
            container.RegisterSingleton<IUserSettingsService, UserSettingsService>();
            container.RegisterSingleton<IPersistentSettingsService, PersistentSettingsService>();
            container.RegisterSingleton<IDialogService, DialogService>();
            container.RegisterSingleton<ICurrentSelectionService, CurrentSelectionService>();
            container.RegisterSingleton<IUnsavedChangesService, UnsavedChangesService>();
            container.RegisterSingleton<IShaftSpeedService, ShaftSpeedService>();
            container.RegisterSingleton<IGraphCursorService, GraphCursorService>();
            container.RegisterSingleton<ISignificantPeaksService, SignificantPeaksService>();
            container.RegisterSingleton<IClipboardService, ClipboardService>();
            container.RegisterSingleton<IReportService, ReportService>();
            container.RegisterSingleton<IApplicationInfo, ApplicationInfo>();
            container.RegisterSingleton(MapperProvider.GetMapper);
            container.RegisterSingleton<IFileSystem, FileSystem>();
            container.RegisterSingleton<VibrationTrendViewModelState>();
            container.RegisterSingleton<TemperatureTrendViewModelState>();
            container.RegisterSingleton<IBearingLibraryRepository, BearingLibraryRepository>();

            return container;
        }

        private void InitializeApplicationSettings()
        {
            var settingsService = _container.GetInstance<IPersistentSettingsService>();
            settingsService.Initialize(LocalStorageFolder);
        }

        private async Task OpenLastUsedDatabaseAsync()
        {
            var databaseService = _container.GetInstance<IDatabaseService>();
            var settingsService = _container.GetInstance<IPersistentSettingsService>();
            var lastUsedDatabase = settingsService.LastUsedDatabase;

            if (lastUsedDatabase == null) return;
            
            try
            {
                await databaseService.OpenAsync(lastUsedDatabase.Name);
            }
            catch (FileNotFoundException)
            {
                Logger.Information($"The database {lastUsedDatabase.Name} at '{lastUsedDatabase.Server}' could not be found");
            }
            catch(Exception e)
            {
                Logger.Warning(e, $"The database {lastUsedDatabase.Name} at '{lastUsedDatabase.Server}' could not be opened");
            }
        }

        private async Task InitializeBearingLibraryAsync()
        {
            var bearingRepo = (BearingLibraryRepository)_container.GetInstance<IBearingLibraryRepository>();
            await bearingRepo.InitializeAsync(Path.Combine(InstallFolder, "BearingLibrary.sqlite"));
        }

        private void ShowMainWindow(CultureInfo currentCulture)
        {
            var mainWindow = _container.GetInstance<MainWindow>();
            mainWindow.AddLanguages(currentCulture, this.OnLanguageSelected);

            var mainWindowViewModel = _container.GetInstance<MainWindowViewModel>();
            mainWindow.DataContext = mainWindowViewModel;
            mainWindow.Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Close any open database connection
            var databaseService = _container.GetInstance<IDatabaseService>();
            databaseService.Dispose();
            
            Logger.Information("Application exit");
            base.OnExit(e);
        }
    }
}
