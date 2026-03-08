using System.IO;
using System.Text;
using System.Windows;

namespace NormSchematicStudio;

public partial class App : Application
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "NormSchematicStudio_startup.log");

    private void Application_Startup(object sender, StartupEventArgs e)
    {
        try
        {
            DispatcherUnhandledException += (_, exArgs) =>
            {
                WriteLog("DispatcherUnhandledException", exArgs.Exception);
                MessageBox.Show(exArgs.Exception.ToString(), "NormSchematicStudio Error", MessageBoxButton.OK, MessageBoxImage.Error);
                exArgs.Handled = true;
            };

            AppDomain.CurrentDomain.UnhandledException += (_, exArgs) =>
            {
                if (exArgs.ExceptionObject is Exception ex)
                    WriteLog("AppDomainUnhandledException", ex);
            };

            WriteLog("Startup", null);

            var w = new MainWindow();
            MainWindow = w;
            w.Show();
            w.Activate();
        }
        catch (Exception ex)
        {
            WriteLog("StartupCatch", ex);
            MessageBox.Show(ex.ToString(), "NormSchematicStudio Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(-1);
        }
    }

    private static void WriteLog(string stage, Exception? ex)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {stage}");
        if (ex != null) sb.AppendLine(ex.ToString());
        sb.AppendLine(new string('-', 60));
        try { Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!); File.AppendAllText(LogPath, sb.ToString(), Encoding.UTF8); } catch { }
    }
}

