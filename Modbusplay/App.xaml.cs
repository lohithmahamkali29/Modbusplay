using System.Windows;
using System.Windows.Threading;

namespace Modbusplay
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(
                $"An error occurred:\n{e.Exception.Message}\n\n" +
                "The application will continue running.\n" +
                "If this keeps happening, try disconnecting and reconnecting.",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            e.Handled = true;
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            e.SetObserved();
            Dispatcher.BeginInvoke(() =>
            {
                MessageBox.Show(
                    $"A background operation failed:\n{e.Exception.InnerException?.Message ?? e.Exception.Message}\n\n" +
                    "If this keeps happening, try disconnecting and reconnecting.",
                    "Background Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            });
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                MessageBox.Show(
                    $"A critical error occurred:\n{ex.Message}\n\n" +
                    "The application may need to be restarted.",
                    "Critical Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}
