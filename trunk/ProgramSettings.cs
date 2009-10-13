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
using System.Xml.Serialization;
using System.IO;
using System.Security;
using System.Windows.Data;
using System.Windows;
using System.ComponentModel;

namespace TieCal
{
    /// <summary>
    /// Holds settings for TieCal
    /// </summary>
    [Serializable]
    public sealed class ProgramSettings : INotifyPropertyChanged
    {
        public ProgramSettings() 
        {
            ReminderMode = ReminderMode.NoReminder;
            ReminderMinutesBeforeStart = 15;
            ConfirmMerge = true;
        }

        private static ProgramSettings _instance;
        public static ProgramSettings Instance
        {
            get 
            {
                if (_instance == null)
                    _instance = LoadSettings();
                return _instance; 
            }
        }
        /// <summary>
        /// Gets the filename where settings are saved.
        /// </summary>
        private static string SaveFilename 
        {
            get
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TieCal");
                Directory.CreateDirectory(folder);
                return Path.Combine(folder, "ProgramSettings.txt");
            }
        }

        public static ProgramSettings LoadSettings()
        {
            try
            {
                if (!File.Exists(SaveFilename))
                    return new ProgramSettings();
                using (TextReader writer = new StreamReader(SaveFilename))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(ProgramSettings));
                    return (ProgramSettings)serializer.Deserialize(writer);
                }
            }
            catch (IOException)
            {
                return new ProgramSettings();
            }
        }

        public void Save()
        {
            if (!RememberPassword)
                NotesPassword = null;
            using (TextWriter writer = new StreamWriter(SaveFilename))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ProgramSettings));
                serializer.Serialize(writer, this);
            }
        }
        private string _notesDb;
        public string NotesDatabase { get { return _notesDb; } set { _notesDb = value; RaisePropertyChanged("NotesDatabase"); } }
        public string NotesPassword { get; set; }
        public bool RememberPassword { get; set; }
        public bool ConfirmMerge { get; set; }
        public bool SyncRepeatingEvents { get; set; }
        private ReminderMode _reminderMode;
        /// <summary>
        /// Gets or sets a value that specifies how reminders should be used for synchronized entries.
        /// </summary>
        public ReminderMode ReminderMode { get { return _reminderMode; } set { _reminderMode = value; RaisePropertyChanged("ReminderMode"); RaisePropertyChanged("ReminderSettingAsString"); } }
        private int _reminderMinutes;
        public int ReminderMinutesBeforeStart { get { return _reminderMinutes; } set { _reminderMinutes = value; RaisePropertyChanged("ReminderMinutesBeforeStart"); RaisePropertyChanged("ReminderSettingAsString"); } }


        public string ReminderSettingAsString
        {
            get
            {
                switch (ReminderMode)
                {
                    case ReminderMode.NoReminder:
                        return "Remove all reminders";
                    case ReminderMode.OutlookDefault:
                        return "Let outlook specify reminder";
                    case ReminderMode.Custom:
                        return String.Format("Remind {0} minutes before meetings", ReminderMinutesBeforeStart);
                    default:
                        return "Unknown setting (" + ReminderMode.ToString() + ")";
                }
            }
        }

        #region INotifyPropertyChanged Members

        private void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }

    public enum ReminderMode
    {
        NoReminder,
        OutlookDefault,
        Custom,
    }

    /// <summary>
    /// Converter that allows an element to be visible if the string value it is bound to is either <c>null</c> or empty (zero length or only whitespaces). 
    /// Useful to display warning icon for invalid strings
    /// </summary>
    public class StringEmptyToVisibilityConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null || value.ToString().Trim().Length == 0)
                return Visibility.Visible;
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

}
