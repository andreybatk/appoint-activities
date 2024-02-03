using System;
using System.Configuration;
using System.Windows;
using System.Windows.Input;
using WpfApplication.Infrastructure.Commands;
using WpfApplication.ViewModels.Base;
using WpfApplication.Views.Windows;

namespace WpfApplication.ViewModels
{
    internal class ConnectionSettingsViewModel : ViewModel
    {
        private string _dataSource;
        private bool _isCheckedBox = true;
        private ConnectionSettings _window;
        public ConnectionSettingsViewModel(ConnectionSettings window)
        {
            _window = window;

            OkCommand = new RelayCommand(OnOkCommandExecuted, CanOkCommandExecute);

            var connection = ConfigurationManager.ConnectionStrings["MsSqlForestEntities"].ConnectionString;
            int found = connection.IndexOf("data source=") + 12;
            int last = connection.IndexOf(';', found);
            DataSource = connection.Substring(found, last - found);
        }
        public bool IsCheckedBox { get => _isCheckedBox; set => Set(ref _isCheckedBox, value); }
        public string DataSource { get => _dataSource; set => Set(ref _dataSource, value); }
        public ICommand OkCommand { get; }

        private bool CanOkCommandExecute(object p)
        {
            if (String.IsNullOrWhiteSpace(DataSource))
            {
                return false;
            }
            return true;
        }
        private void OnOkCommandExecuted(object p)
        {
            try
            {
                var connection = ConfigurationManager.ConnectionStrings["MsSqlForestEntities"].ConnectionString;
                int found = connection.IndexOf("data source=") + 12;
                int last = connection.IndexOf(';', found);
                connection = connection.Remove(found, last - found);
                connection = connection.Insert(found, DataSource);

                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var connectionStringsSection = (ConnectionStringsSection)config.GetSection("connectionStrings");

                connectionStringsSection.ConnectionStrings["MsSqlForestEntities"].ConnectionString = connection;
                config.Save();
                ConfigurationManager.RefreshSection("connectionStrings");

                _window.DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при изменении строки подключения: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
    }
}
