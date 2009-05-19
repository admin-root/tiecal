// Part of the TieCal project (http://code.google.com/p/tiecal/)
// Copyright (C) 2009, Isak Savo <isak.savo@gmail.com>
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//      http://www.gnu.org/licenses/gpl.html
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
using System.Globalization;

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
            this.DataContext = modifiedEntries;
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
        private CultureInfo enUS;
        public EntryToDurationConverter()
        {
            enUS = CultureInfo.CreateSpecificCulture("en-US");
        }
        private string PluralEnding(int number)
        {
            if (number == 1)
                return "";
            return "s";
        }

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            CalendarEntry entry = (CalendarEntry) value;
            var duration = entry.EndTime - entry.StartTime;
            var friendlyDateFormat = "ddd, MMM dd";
            if (entry.IsAllDay)
            {
                int days = (int) Math.Round(duration.TotalDays, 0);
                if (days == 1)
                    return String.Format("{0} (all day)", entry.StartTime.ToString(friendlyDateFormat, enUS));
                else
                    // Not sure if this can actually happen, I think they will be two separate entries...
                    return String.Format("{0:d} — {1:d} ({2} day{3})", entry.StartTime, entry.EndTime, days, PluralEnding(days));
            }
            else if (entry.StartTime.Date == entry.EndTime.Date)
            {
                return String.Format("{0} {1:t} ({2:0.#} hours)", entry.StartTime.ToString(friendlyDateFormat, enUS), entry.StartTime, duration.TotalHours);
            }
            else
            {
                // Starts and ends on different days
                StringBuilder sb = new StringBuilder(entry.StartTime.ToString(friendlyDateFormat, enUS));
                sb.AppendFormat(" {0:t}", entry.StartTime);
                sb.Append(" — ");
                sb.AppendFormat("{0} {1:t}", entry.EndTime.ToString("MMM dd", enUS), entry.EndTime);
                sb.Append(" (");
                if (duration.Days > 0)
                    sb.AppendFormat("{0} day{1},", duration.Days, PluralEnding(duration.Days));
                sb.AppendFormat("{0:0.#} hours", duration.TotalHours - (24 * duration.Days));
                sb.Append(")");
                return sb.ToString();
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

    internal class ModificationTypeToImageConverter : IValueConverter
    {
        private static ImageSource[] sources = null;

        public ModificationTypeToImageConverter()
        {
            if (sources == null)
            {
                sources = new ImageSource[3];
                sources[(int)Modification.New] = new BitmapImage(new Uri("pack://application:,,,/Images/Add-64.png", UriKind.Absolute));
                sources[(int)Modification.Modified] = new BitmapImage(new Uri("pack://application:,,,/Images/Edit-64.png", UriKind.Absolute));
                sources[(int)Modification.Removed] = new BitmapImage(new Uri("pack://application:,,,/Images/trash-64.png", UriKind.Absolute));
            }
        }
        #region IValueConverter Members
        
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            Modification mod = (Modification)value;
            return sources[(int)mod];
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

}
