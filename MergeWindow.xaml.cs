using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Collections;

namespace TieCal
{
    /// <summary>
    /// Interaction logic for MergeWindow.xaml
    /// </summary>
    public partial class MergeWindow : Window
    {
        public MergeWindow(IEnumerable<CalendarEntry> newEntries, IEnumerable changedEntries, IEnumerable<CalendarEntry> oldEntries)
        {
            InitializeComponent();
            lstNewEntries.ItemsSource = newEntries;
            lstModifiedEntries.ItemsSource = changedEntries;
            lstOldEntries.ItemsSource = oldEntries;
        }

        private void btnMerge_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            Close();
        }
    }
}
