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
        public MergeWindow()
        {
            InitializeComponent();
        }

        public MergeWindow(IEnumerable<ModifiedEntry> modifiedEntries)
            : this()
        {
            lstModifiedEntries.ItemsSource = modifiedEntries;
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

    /// <summary>
    /// Represents the different types of updates to a calendar entry.
    /// </summary>
    public enum Modification
    {
        /// <summary>The entry has never been synced before</summary>
        New,
        /// <summary>The entry has been synced before, but some fields has changed</summary>
        Modified,
        /// <summary>The entry was synced before, but does no longer exist in Lotus Notes</summary>
        Removed
    }

    public class ModifiedEntry
    {
        public ModifiedEntry(CalendarEntry entry, Modification modification)
        {
            Entry = entry;
            Modification = modification;
            ApplyModification = true;
        }

        public ModifiedEntry(CalendarEntry entry, Modification modification, IEnumerable<string> changedFields)
            : this(entry, modification)
        {
            if (modification != Modification.Modified)
                throw new InvalidOperationException("This constructor is only valid when modification is Modification.Modified");
            ChangedFields = changedFields;
        }

        public bool ApplyModification { get; set; }
        public Modification Modification { get; set; }
        public CalendarEntry Entry { get; set; }
        public IEnumerable<string> ChangedFields { get; set; }
    }

    internal class EntryToDurationConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            CalendarEntry entry = (CalendarEntry) value;
            var duration = entry.EndTime - entry.StartTime;
            if (entry.StartTime.Date == entry.EndTime.Date)
            {
                return String.Format("{0:d} {1:t}-{2:t} ({3:0.#} hours)", entry.StartTime, entry.StartTime, entry.EndTime, duration.TotalHours);
            }
            else
            {
                return String.Format("{0:g}-{1:g} ({2} days, {3:0.#} hours)", entry.StartTime, entry.EndTime, duration.Days, duration.TotalHours - (24 * duration.Days));
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

}
