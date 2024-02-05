using System;
using System.Configuration;
using System.Text;
using System.Web;
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
        private string _customConnectionString;
        private bool _isCheckedBoxCustomConnection = false;
        private bool _isEnableTextBoxCustomConnection = false;
        private bool _isEnableTextBox = true;
        private ConnectionSettings _window;

        public ConnectionSettingsViewModel(ConnectionSettings window)
        {
            _window = window;

            OkCommand = new RelayCommand(OnOkCommandExecuted, CanOkCommandExecute);

            SetTextBoxInfo();
        }
        public bool IsCheckedBoxCustomConnection
        {
            get
            {
                return _isCheckedBoxCustomConnection;
            }
            set
            {
                Set(ref _isCheckedBoxCustomConnection, value);
                IsEnableTextBoxCustomConnection = value;
                IsEnableTextBox = !value;
            }
        }
        public bool IsEnableTextBoxCustomConnection { get => _isEnableTextBoxCustomConnection; set => Set(ref _isEnableTextBoxCustomConnection, value); }
        public bool IsEnableTextBox { get => _isEnableTextBox; set => Set(ref _isEnableTextBox, value); }
        public string DataSource { get => _dataSource; set => Set(ref _dataSource, value); }
        public string CustomConnectionString { get => _customConnectionString; set => Set(ref _customConnectionString, value); }
        public string InitialCatalog { get => _initialCatalog; set => Set(ref _initialCatalog, value); }
        public ICommand OkCommand { get; }

        private bool CanOkCommandExecute(object p)
        {
            if ((!String.IsNullOrEmpty(DataSource) && !String.IsNullOrEmpty(InitialCatalog)) || !String.IsNullOrEmpty(CustomConnectionString))
            {
                return true;
            }
            return false;
        }
        private void OnOkCommandExecuted(object p)
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var connectionStringsSection = (ConnectionStringsSection)config.GetSection("connectionStrings");
                string con = "";

                if(IsCheckedBoxCustomConnection)
                {
                    con = new String($"metadata=res://*/ModelDB.csdl|res://*/ModelDB.ssdl|res://*/ModelDB.msl;provider=System.Data.SqlClient;provider connection string=&quot;{CustomConnectionString};App=EntityFramework&quot;");
                }
                else
                {
                    con = new String($"metadata=res://*/ModelDB.csdl|res://*/ModelDB.ssdl|res://*/ModelDB.msl;provider=System.Data.SqlClient;provider connection string=&quot;Data Source={DataSource};Initial Catalog={InitialCatalog};integrated security=True;MultipleActiveResultSets=True;App=EntityFramework&quot;");
                

                }

                string myEncodedString = HttpUtility.HtmlDecode(con);
                connectionStringsSection.ConnectionStrings["MsSqlForestEntities"].ConnectionString = myEncodedString;

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
            var connection = ConfigurationManager.ConnectionStrings["MsSqlForestEntities"].ConnectionString;
            
            int found = connection.IndexOf("data source=", StringComparison.CurrentCultureIgnoreCase);
            if (found != -1)
            {
                found += 12;
                int last = connection.IndexOf(';', found);
                DataSource = connection.Substring(found, last - found);
            }

            found = connection.IndexOf("initial catalog=", StringComparison.CurrentCultureIgnoreCase);
            if (found != -1)
            {
                found += 16;
                int last = connection.IndexOf(';', found);
                InitialCatalog = connection.Substring(found, last - found);
            }
        }
    }
}
