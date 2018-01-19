using System.Windows;
using Cassini.UI.ViewModel;
using MahApps.Metro.Controls;

namespace Cassini.UI
{

    public partial class MainWindow : MetroWindow
    {
        private MainViewModel _mainViewModel;

        public MainWindow(MainViewModel mainViewModel)
        {
            InitializeComponent();
            _mainViewModel = mainViewModel;
            DataContext = _mainViewModel;
            Loaded += MainViewLoaded;
        }

        private async void MainViewLoaded(object sender, RoutedEventArgs e)
        {
            await _mainViewModel.LoadAsync();
        }
    }
}
