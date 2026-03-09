using System.Collections.Specialized;
using System.Windows;
using Modbusplay.ViewModels;

namespace Modbusplay
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            var vm = new MainViewModel();
            DataContext = vm;
            vm.FrameLog.CollectionChanged += FrameLog_CollectionChanged;
        }

        private void FrameLog_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (LogListBox.Items.Count > 0)
                LogListBox.ScrollIntoView(LogListBox.Items[LogListBox.Items.Count - 1]);
        }

        protected override void OnClosed(EventArgs e)
        {
            (DataContext as IDisposable)?.Dispose();
            base.OnClosed(e);
        }
    }
}