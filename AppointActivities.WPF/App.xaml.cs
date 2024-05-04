using System.Threading;
using System.Windows;

namespace AppointActivities.WPF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static Mutex _syncObject;
        private const string _syncObjectName = "{E663FA11-AE0D-480e-9FCA-4BE9B8CDB4E91}";

        public App()
        {
            bool createdNew;
            _syncObject = new Mutex(true, _syncObjectName, out createdNew);
            if (!createdNew)
            {
                MessageBox.Show("Программа уже запущена.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
    }
}
