using System;
using System.ComponentModel;
using System.Data.Entity;
using System.Windows;
using System.Windows.Input;
using WpfApplication.DB;
using WpfApplication.Infrastructure;
using WpfApplication.Infrastructure.Commands;
using WpfApplication.Models;
using WpfApplication.ViewModels.Base;
using WpfApplication.Views.Windows;

namespace WpfApplication.ViewModels
{
    internal class MainWindowViewModel : ViewModel
    {
        private BindingList<MyTable> _data;
        private MsSqlForestEntities dbContext;
        private string _currentActivitieInfo;  

        public MainWindowViewModel()
        {
            dbContext = new MsSqlForestEntities();
            Preparing();

            ChangeConectionCommand = new RelayCommand(OnChangeConectionCommandExecuted);
            UpdataDBCommand = new RelayCommand(OnUpdataDBCommandExecuted);
            ActivitieCommand = new RelayCommand(OnActivitieCommandExecuted);
            Activitie2Command = new RelayCommand(OnActivitie2CommandExecuted);
            Activitie3Command = new RelayCommand(OnActivitie3CommandExecuted);
            Activitie4Command = new RelayCommand(OnActivitie4CommandExecuted);
            Activitie5Command = new RelayCommand(OnActivitie5CommandExecuted);
            Activitie6Command = new RelayCommand(OnActivitie6CommandExecuted);
        }

        public BindingList<MyTable> DataList { get => _data; set => Set(ref _data, value); }
        public string CurrentActivitieInfo { get => _currentActivitieInfo; set => Set(ref _currentActivitieInfo, value); }
        
        public ICommand ChangeConectionCommand { get; }
        public ICommand UpdataDBCommand { get; }
        public ICommand ActivitieCommand { get; }
        public ICommand Activitie2Command { get; }
        public ICommand Activitie3Command { get; }
        public ICommand Activitie4Command { get; }
        public ICommand Activitie5Command { get; }
        public ICommand Activitie6Command { get; }

        private void Preparing()
        {
            try
            {
                dbContext.MyTable.Load();
                DataList = dbContext.MyTable.Local.ToBindingList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке базы данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            } 
        }

        private void OnUpdataDBCommandExecuted(object p)
        {
            try
            {
                DataList = dbContext.MyTable.Local.ToBindingList();
                MessageBox.Show("Таблица обновлена.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnActivitieCommandExecuted(object p)
        {
            try
            {
                foreach (var item in DataList)
                {
                    IActivitie appointActivitie = new Activitiy(item);
                    appointActivitie.CalculateActivitie();
                    appointActivitie.AppointActivitie();
                }
                CurrentActivitieInfo = "Текущий сценарий: 1";
                dbContext.SaveChanges();
                MessageBox.Show("1 сценарий завершен успешно.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnActivitie2CommandExecuted(object p)
        {
            try
            {
                foreach (var item in DataList)
                {
                    IActivitie appointActivitie = new Activitiy2(item);
                    appointActivitie.CalculateActivitie();
                    appointActivitie.AppointActivitie();
                }
                CurrentActivitieInfo = "Текущий сценарий: 2";
                dbContext.SaveChanges();
                MessageBox.Show("2 сценарий завершен успешно.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnActivitie3CommandExecuted(object p)
        {
            try
            {
                foreach (var item in DataList)
                {
                    IActivitie appointActivitie = new Activitiy3(item);
                    appointActivitie.CalculateActivitie();
                    appointActivitie.AppointActivitie();
                }
                CurrentActivitieInfo = "Текущий сценарий: 3";
                dbContext.SaveChanges();
                MessageBox.Show("3 сценарий завершен успешно.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnActivitie4CommandExecuted(object p)
        {
            try
            {
                foreach (var item in DataList)
                {
                    IActivitie appointActivitie = new Activitiy4(item);
                    appointActivitie.CalculateActivitie();
                    appointActivitie.AppointActivitie();
                }
                CurrentActivitieInfo = "Текущий сценарий: 4";
                dbContext.SaveChanges();
                MessageBox.Show("4 сценарий завершен успешно.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnActivitie5CommandExecuted(object p)
        {
            try
            {
                foreach (var item in DataList)
                {
                    IActivitie appointActivitie = new Activitiy5(item);
                    appointActivitie.CalculateActivitie();
                    appointActivitie.AppointActivitie();
                }
                CurrentActivitieInfo = "Текущий сценарий: 5";
                dbContext.SaveChanges();
                MessageBox.Show("5 сценарий завершен успешно.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnActivitie6CommandExecuted(object p)
        {
            try
            {
                foreach (var item in DataList)
                {
                    IActivitie appointActivitie = new Activitiy6(item);
                    appointActivitie.CalculateActivitie();
                    appointActivitie.AppointActivitie();
                }
                CurrentActivitieInfo = "Текущий сценарий: 6";
                dbContext.SaveChanges();
                MessageBox.Show("6 сценарий завершен успешно.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnChangeConectionCommandExecuted(object p)
        {
            ConnectionSettings window = new ConnectionSettings();
            ConnectionSettingsViewModel windowViewModel = new ConnectionSettingsViewModel(window);
            window.DataContext = windowViewModel;
            window.ShowDialog();

            if(window.DialogResult.Value)
            {
                if(windowViewModel.IsCheckedBox)
                {
                    Application.Current.Shutdown();
                }
            }
        }
    }
}
