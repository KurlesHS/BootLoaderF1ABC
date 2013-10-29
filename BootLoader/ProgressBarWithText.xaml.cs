using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace BootLoader
{
    public class RectConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var width = (double)values[0];
            var height = (double)values[1];
            return new Rect(0, 0, width, height);
        }
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    
    public partial class ProgressBarWithText : ProgressBar
    {


        public ProgressBarWithText()
        {
            Text = "";
            InitializeComponent();
        }


        public string Text
        {
            get { return (string)GetValue(TextProperty); }
            set { SetValue(TextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Text.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty TextProperty =
            DependencyProperty.Register("Text", typeof(string), typeof(ProgressBarWithText), new PropertyMetadata(""));


    }
}
