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

namespace TieCal
{
    /// <summary>
    /// Holds settings for TieCal
    /// </summary>
    [Serializable]
    public sealed class ProgramSettings
    {
        public ProgramSettings() 
        {
            ReminderMode = ReminderMode.NoReminder;
            ReminderMinutesBeforeStart = 15;
        }

        private static ProgramSettings _instance = ProgramSettings.LoadSettings();
        public static ProgramSettings Instance
        {
            get { return _instance; }
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

        public string NotesDatabase { get; set; }
        public string NotesPassword { get; set; }
        public bool RememberPassword { get; set; }
        public ReminderMode ReminderMode { get; set; }
        public int ReminderMinutesBeforeStart { get; set; }
    }

    public enum ReminderMode
    {
        NoReminder,
        OutlookDefault,
        Custom,
    }
}
