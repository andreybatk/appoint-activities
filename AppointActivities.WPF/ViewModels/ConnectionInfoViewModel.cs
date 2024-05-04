using System.Windows.Input;
using AppointActivities.WPF.Infrastructure.Commands;
using AppointActivities.WPF.ViewModels.Base;
using AppointActivities.WPF.Views.Windows;

namespace AppointActivities.WPF.ViewModels
{
    internal class ConnectionInfoViewModel : ViewModel
    {
        private string _info;
        private ConnectionInfo _window;

        public ConnectionInfoViewModel(ConnectionInfo window)
        {
            _window = window;

            OkCommand = new RelayCommand(OnOkCommandExecuted);

            SetTextBoxInfo();
        }

        public string Info { get => _info; set => Set(ref _info, value); }
        public ICommand OkCommand { get; }

        private void OnOkCommandExecuted(object p)
        {
            _window.DialogResult = true;
        }
        private void SetTextBoxInfo()
        {
            Info = @"select
                'data source=' + @@servername +
                ';initial catalog=' + db_name() +
                case type_desc
                    when 'WINDOWS_LOGIN' 
                        then ';trusted_connection=true'
                    else
                        ';user id=' + suser_name() + ';password=<<YourPassword>>'
                end
                as ConnectionString
            from sys.server_principals
            where name = suser_name()";
        }
    }
}