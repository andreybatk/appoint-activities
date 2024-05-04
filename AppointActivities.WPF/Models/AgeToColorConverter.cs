using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace AppointActivities.WPF.Models
{
    public class AgeToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (int)value <= 16 ?
                new SolidColorBrush(Colors.OrangeRed)
                : new SolidColorBrush(Colors.White);
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("The method or operation is not implemented.");
        }
    }
}
