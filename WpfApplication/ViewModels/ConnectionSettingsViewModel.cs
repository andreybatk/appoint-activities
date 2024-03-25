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
        private string _initialCatalog;
        private bool _isEnableTextBoxCustomConnection = false;
        private bool _isEnableTextBox = true;
        private ConnectionSettings _window;

        public ConnectionSettingsViewModel(ConnectionSettings window)
        {
            _window = window;

            OkCommand = new RelayCommand(OnOkCommandExecuted, CanOkCommandExecute);
            FilePathCommand = new RelayCommand(OnFilePathCommandExecuted);
            SetTextBoxInfo();
        }

        public bool IsEnableTextBoxCustomConnection { get => _isEnableTextBoxCustomConnection; set => Set(ref _isEnableTextBoxCustomConnection, value); }
        public bool IsEnableTextBox { get => _isEnableTextBox; set => Set(ref _isEnableTextBox, value); }
        public string DataSource { get => _dataSource; set => Set(ref _dataSource, value); }
        public string InitialCatalog { get => _initialCatalog; set => Set(ref _initialCatalog, value); }
        public ICommand OkCommand { get; }
        public ICommand FilePathCommand { get; }

        private bool CanOkCommandExecute(object p)
        {
            if ((!String.IsNullOrEmpty(DataSource) && !String.IsNullOrEmpty(InitialCatalog)))
            {
                return true;
            }
            return false;
        }
        private void OnFilePathCommandExecuted(object p)
        {
            try
            {
                var dialog = new Microsoft.Win32.OpenFileDialog();
                dialog.FileName = "BD_AIS_POL";
                dialog.DefaultExt = ".mdf";
                dialog.Filter = "MSSQL (.mdf)|*.mdf";

                bool? result = dialog.ShowDialog();

                if (result == true)
                {
                    InitialCatalog = dialog.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void OnOkCommandExecuted(object p)
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var connectionStringsSection = (ConnectionStringsSection)config.GetSection("connectionStrings");

                var con = new String($"metadata = res://*/ModelAISPOL.csdl|res://*/ModelAISPOL.ssdl|res://*/ModelAISPOL.msl;provider=System.Data.SqlClient;provider connection string=\"Data Source={DataSource};AttachDbFilename={InitialCatalog};Integrated Security=True;Connect Timeout=30;App=EntityFramework\"");

                connectionStringsSection.ConnectionStrings["BD_AIS_POLEntities"].ConnectionString = con;

                config.Save();
                ConfigurationManager.RefreshSection("connectionStrings");

                MessageBox.Show("Для сохранения изменений приложение закроется.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);

                _window.DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при изменении строки подключения: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void SetTextBoxInfo()
        {
            var connection = ConfigurationManager.ConnectionStrings["BD_AIS_POLEntities"].ConnectionString;

            int found = connection.IndexOf("data source=", StringComparison.CurrentCultureIgnoreCase);
            if (found != -1)
            {
                found += 12;
                int last = connection.IndexOf(';', found);
                DataSource = connection.Substring(found, last - found);
            }

            found = connection.IndexOf("AttachDbFilename=", StringComparison.CurrentCultureIgnoreCase);
            if (found != -1)
            {
                found += 17;
                int last = connection.IndexOf(';', found);
                InitialCatalog = connection.Substring(found, last - found);
            }
        }
    }
}