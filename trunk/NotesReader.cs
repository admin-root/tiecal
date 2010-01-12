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
using Domino;
using System.Diagnostics;
using System.ComponentModel;
using System.Security;
using System.Runtime.InteropServices;

namespace TieCal
{
    public class NotesReader : ICalendarReader
    {
        private List<CalendarEntry> _calendarEntries;
        public NotesReader()
        {
            FetchCalendarWorker = new BackgroundWorker();
            FetchCalendarWorker.WorkerReportsProgress = true;
            FetchCalendarWorker.WorkerSupportsCancellation = true;
            FetchCalendarWorker.DoWork += new DoWorkEventHandler(worker_DoWork);
        }
        
        private ISession CreateNotesSession()
        {
            var session = new NotesSessionClass();
            session.Initialize(ProgramSettings.Instance.NotesPassword);
            CalendarAppVersion = session.NotesVersion;
            return session;
        }

        public List<string> GetAvailableDatabases()
        {
            List<String> databases = new List<string>();
            var session = CreateNotesSession();            
            var dir = session.GetDbDirectory("");
            var db = dir.GetFirstDatabase(DB_TYPES.DATABASE);
            while (db != null)
            {
                databases.Add(db.FilePath);
                db = dir.GetNextDatabase();
            }
            
            return databases;
        }

        /// <summary>
        /// Internal helper class for parsing the text and date data received from the Notes COM object
        /// </summary>
        private class NotesEntryItems
        {
            private Dictionary<string, string> stringItems;
            private Dictionary<string, DateTime> dateItems;
            public TimeSpan StartTimeZoneOffset { get; private set; }
            public TimeSpan EndTimeZoneOffset { get; private set; }
            public TimeSpan LocalTimeZoneOffset { get; private set; }

            /// <summary>
            /// Gets the list of occurrences for this calendar entry.
            /// </summary>
            public List<DateTime> Occurrences { get; private set; }
            /// <summary>
            /// Initializes a new instance of the <see cref="NotesEntryItems"/> class based on the data in the provided notes entry.
            /// </summary>
            /// <param name="notesEntry">The notes entry to read data from.</param>
            public NotesEntryItems(NotesViewEntry notesEntry)
            {
                stringItems = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
                dateItems = new Dictionary<string, DateTime>(StringComparer.InvariantCultureIgnoreCase);
                Occurrences = new List<DateTime>();
                StartTimeZoneOffset = TimeSpan.FromTicks(0);
                EndTimeZoneOffset = TimeSpan.FromTicks(0);
                LocalTimeZoneOffset = TimeSpan.FromTicks(0);
                var items = (object[])notesEntry.Document.Items;

                for (int i = 0; i < items.Length; i++)
                {
                    NotesItem item = (NotesItem)items[i];
                    if (stringItems.ContainsKey(item.Name))
                        // Ignore duplicate items ('Received' is for instance specified multiple times)
                        continue;
                    if (item.Text != null)
                        stringItems.Add(item.Name, item.Text);
                    if (item.DateTimeValue != null)
                    {
                        dateItems.Add(item.Name, (DateTime)item.DateTimeValue.LSGMTTime);
                        if (item.Name == "StartTime" || item.Name == "EndTime")
                            // We need the local time for some all-day events in case they don't have time-zone info
                            dateItems.Add(item.Name + "-local", (DateTime)item.DateTimeValue.LSLocalTime);
                    }
                    if (item.Name == "CalendarDateTime")
                    {
                        // This is a list of all occurrences (including original one)
                        object times = item.GetValueDateTimeArray();
                        foreach (object time in (object[])times)
                        {
                            var nTime = time as NotesDateTime;
                            Occurrences.Add((DateTime)nTime.LSGMTTime);
                        }
                    }
                    else if (item.Name == "StartTimeZone")
                    {
                        StartTimeZoneOffset = GetTimeZoneDiff(item);
                    }
                    else if (item.Name == "EndTimeZone")
                    {
                        EndTimeZoneOffset = GetTimeZoneDiff(item);
                    }
                    else if (item.Name == "LocalTimeZone")
                    {
                        LocalTimeZoneOffset = GetTimeZoneDiff(item);
                    }
                }

                if (dateItems.ContainsKey("StartDateTime"))
                    StartTime = dateItems["StartDateTime"];
                else if (dateItems.ContainsKey("StartDate"))
                {
                    StartTime = dateItems["StartDate"];
                    if (dateItems.ContainsKey("StartTime"))
                    {
                        var timePart = dateItems["StartTime"].TimeOfDay;
                        StartTime = new DateTime(StartTime.Year, StartTime.Month, StartTime.Day, timePart.Hours, timePart.Minutes, timePart.Seconds);
                    }
                    StartTime = StartTime.Add(StartTimeZoneOffset);
                }
                else if (dateItems.ContainsKey("StartTime"))
                {
                    Debugger.Break(); // This is probably an exception
                    StartTime = dateItems["StartTime"].Add(StartTimeZoneOffset);
                    Debug.Assert(StartTime.Date != default(DateTime).Date);
                }

                if (dateItems.ContainsKey("EndDateTime"))
                    EndTime = dateItems["EndDateTime"];
                else if (dateItems.ContainsKey("EndDate"))
                {
                    EndTime = dateItems["EndDate"];
                    if (dateItems.ContainsKey("EndTime"))
                    {
                        var timePart = dateItems["EndTime"].TimeOfDay;
                        EndTime = EndTime = new DateTime(EndTime.Year, EndTime.Month, EndTime.Day,
                            timePart.Hours, timePart.Minutes, timePart.Seconds);
                    }
                    EndTime = EndTime.Add(EndTimeZoneOffset);
                }
                else if (dateItems.ContainsKey("EndTime"))
                {
                    Debugger.Break();
                    EndTime = dateItems["EndTime"].Add(EndTimeZoneOffset);
                    Debug.Assert(EndTime.Date != default(DateTime).Date);
                }
            }

            /// <summary>
            /// Gets the start time, in UTC
            /// </summary>
            /// <value>The start time.</value>
            public DateTime StartTime { get; private set; }
            /// <summary>
            /// Gets the end time, in UTC.
            /// </summary>
            public DateTime EndTime { get; private set; }

            /// <summary>
            /// Determines whether there is a date item associated with the specified key.
            /// </summary>
            /// <param name="key">The key to look for.</param>
            /// <returns>
            /// 	<c>true</c> if the specified key exists; otherwise, <c>false</c>.
            /// </returns>
            public bool HasDateItem(string key)
            {
                return dateItems.ContainsKey(key);
            }

            /// <summary>
            /// Determines whether there is a string item associated with the specified key.
            /// </summary>
            /// <param name="key">The key to look for.</param>
            /// <returns>
            /// 	<c>true</c> if the specified key exists; otherwise, <c>false</c>.
            /// </returns>
            public bool HasStringItem(string key)
            {
                return stringItems.ContainsKey(key);
            }

            /// <summary>
            /// Gets the date item for the specified key, or an "empty" datetime if the key doesn't exist.
            /// </summary>
            /// <param name="key">The key to look for</param>
            public DateTime GetDateItemOrDefault(string key)
            {
                if (!dateItems.ContainsKey(key))
                    return default(DateTime);
                return dateItems[key];
            }

            /// <summary>
            /// Gets the string item for the specified key, or <c>null</c> if the key doesn't exist
            /// </summary>
            /// <param name="key">The key to look for</param>
            public string GetStringItemOrDefault(string key)
            {
                if (!stringItems.ContainsKey(key))
                    return null;
                return stringItems[key];
            }
        }
        /// <summary>
        /// Creates the calendar entry from the provided Lotus Notes calendar entry.
        /// </summary>
        /// <param name="notesEntry">The notes entry.</param>
        /// <remarks>
        /// More details about Notes API: http://www-01.ibm.com/support/docview.wss?rs=463context=SSKTMJ&amp;context=SSKTWP&amp;dc=DB520&amp;dc=D600&amp;dc=DB530&amp;dc=D700&amp;dc=DB500&amp;dc=DB540&amp;dc=DB510&amp;dc=DB550&amp;q1=1229486&amp;uid=swg21229486&amp;loc=en_US&amp;cs=utf-8&amp;lang=en
        /// </remarks>
        /// <returns></returns>
        private static CalendarEntry CreateCalendarEntry(NotesViewEntry notesEntry, out SkippedEntry skippedEntry)
        {
            CalendarEntry newEntry = new CalendarEntry();
            newEntry.NotesID = notesEntry.UniversalID;
            NotesDocument doc = notesEntry.Document;
            Debug.Assert(doc.UniversalID == notesEntry.UniversalID);
            var items = new NotesEntryItems(notesEntry);
            skippedEntry = null;
            newEntry.Body = items.GetStringItemOrDefault("Body");
            newEntry.Subject = items.GetStringItemOrDefault("Subject");
            if (String.IsNullOrEmpty(newEntry.Subject))
                newEntry.Subject = "(no subject)";

            if (items.HasStringItem("Location"))
                newEntry.Location = GetNameFromNotesName(items.GetStringItemOrDefault("Location"));
            if (items.HasStringItem("Room"))
            {
                var room = GetNameFromNotesName(items.GetStringItemOrDefault("Room"));
                if (!String.IsNullOrEmpty(newEntry.Location))
                    newEntry.Location = String.Format("{0}, {1}", newEntry.Location, room);
                else
                    newEntry.Location = room;
            }
            var people = items.GetStringItemOrDefault("SendTo");
            if (people != null)
                newEntry.Participants.AddRange(GetRecipentList(people));
            people = items.GetStringItemOrDefault("CopyTo");
            if (people != null)
                newEntry.OptionalParticipants.AddRange(GetRecipentList(people));
            newEntry.From = items.GetStringItemOrDefault("From");
            newEntry.StartTime = items.StartTime;
            newEntry.EndTime = items.EndTime;
            
            // Notes stores timezone offset in the inverted form (e.g. CET would be -1, instead of +1)
            newEntry.SetEndTimeZoneFromOffset(TimeSpan.FromTicks(items.EndTimeZoneOffset.Ticks * -1));
            newEntry.SetStartTimeZoneFromOffset(TimeSpan.FromTicks(items.StartTimeZoneOffset.Ticks * -1));
            var appointmentType = items.GetStringItemOrDefault("AppointmentType");
            
            // Sanity check
            if (items.HasStringItem("TaskType") && appointmentType == null)
            {
                // It's probably a TODO or Followup, ignore it
                skippedEntry = new SkippedEntry(newEntry, "Not a valid calendar entry. Could be TODO or Followup item");
                return null;
            }
            //if (newEntry.Subject.IndexOf("Nyårsafton", StringComparison.CurrentCultureIgnoreCase) >= 0)
            //    Debugger.Break();

            if (appointmentType == "2")
            {
                /* All Day Event */
                newEntry.IsAllDay = true;
                newEntry.StartTimeZone = TimeZoneInfo.Local;
                newEntry.EndTimeZone = TimeZoneInfo.Local;
                // Reset the time to zero to reflect that it is all day event
                newEntry.StartTime = newEntry.StartTime.Date;
                newEntry.EndTime = newEntry.EndTime.Date.AddDays(1);
            }
            else if (appointmentType == "1")
            {
                // Anniversary
                newEntry.StartTime = items.GetDateItemOrDefault("StartTime-local");
                newEntry.EndTime = items.GetDateItemOrDefault("EndTime-local");
                newEntry.StartTimeZone = TimeZoneInfo.Local;
                newEntry.EndTimeZone = TimeZoneInfo.Local;
                // Reset the time to zero to reflect that it is all day event
                newEntry.StartTime = newEntry.StartTime.Date;
                newEntry.EndTime = newEntry.EndTime.Date.AddDays(1);
                newEntry.IsAllDay = true;
            }
            if (!items.HasStringItem("OrgRepeat") && items.Occurrences.Count > 1)
            {
                // OrgRepeat should be set on all repeating events, just treat it as a normal event instead
                items.Occurrences.Clear();
                items.Occurrences.Add(newEntry.StartTime);
            }
            if (items.Occurrences.Count > 1)
            {
                try
                {
                    if (newEntry.Subject.StartsWith("test-"))
                        Debugger.Break();
                    newEntry.SetRepeatPattern(items.Occurrences);
                }
                catch (ArgumentException)
                {
                    skippedEntry = new SkippedEntry(newEntry, "Could not find a valid repeat pattern for recurring event.");
                    return null;
                }
            }
            
            Debug.Assert(items.Occurrences.Count > 0);
            Debug.Assert(newEntry.NotesID != null && newEntry.NotesID.Length > 0);

            return newEntry;
        }

        /// <summary>
        /// Gets a value indicating whether this notes reader has access to notes.
        /// </summary>
        public bool HasAccessToNotes
        {
            get { return !String.IsNullOrEmpty(ProgramSettings.Instance.NotesPassword); }
        }
        /// <summary>
        /// Converts a list of people in a notes string format to an enumerable list of display friendly names.
        /// </summary>
        /// <param name="notesArray">The notes array, with each item separated by semicolon.</param>
        private static IEnumerable<string> GetRecipentList(string notesArray)
        {
            var people = SplitNotesArray(notesArray);
            for (int i = 0; i < people.Count; i++)
            {
                people[i] = GetNameFromNotesName(people[i]);
            }
            return people;
        }

        private static TimeSpan GetTimeZoneDiff(NotesItem item)
        {
            NotesDateTime dt = item.DateTimeValue;
            if (dt != null)
                return TimeSpan.FromHours(dt.TimeZone);
            // Parse string: 
            // Z=-1$DO=1$DL=3 -1 1 10 -1 1$ZX=87$ZN=W. Europe
            // Z=5$DO=1$DL=3 2 1 11 1 1$ZX=28$ZN=Eastern
            // Z=-3005$DO=0$ZX=35$ZN=India
            // Z=-3004$DO=0$ZX=0$ZN=Afghanistan
            int start = item.Text.IndexOf("Z=");
            if (start >= 0)
                start += "Z=".Length;
            int end = item.Text.IndexOf("$", start);
            if (end == -1)
                end = item.Text.Length;
            string strDiff = item.Text.Substring(start, end - start);
            int diff;
            if (!int.TryParse(strDiff, out diff))
                return new TimeSpan(0);
            if (Math.Abs(diff) <= 12)
                return TimeSpan.FromHours(diff);
            else if (Math.Abs(diff) < 1000)
                Debugger.Break();
            // Timezones with 30 minute offsets are stored as MMHH
            int hour = diff % 100; 
            int minute = diff / 100;

            return new TimeSpan(hour, minute, 0);
            
        }
        /// <summary>
        /// Gets the name part of a lotus notes name field (CN=The Name/OU=etc.).
        /// </summary>
        /// <param name="notesName">The name field, from Lotus Notes.</param>
        /// <returns></returns>
        private static string GetNameFromNotesName(string notesName)
        {
            int start = notesName.IndexOf("CN=");
            if (start == -1)
                return notesName;
            else
                start += "CN=".Length;
            int end = notesName.IndexOf("/", start);
            if (end != -1)
                return notesName.Substring(start, end - start);
            else
                return notesName.Substring(start);
        }

        private static List<string> SplitNotesArray(string notesArray)
        {
            if (notesArray == null)
                throw new ArgumentNullException("notesArray", "Cannot split a NULL string");
            List<string> items = new List<string>();
            if (notesArray.IndexOf(';') != -1)
            {
                items.AddRange(notesArray.Split(';'));
            }
            else
                items.Add(notesArray);
            return items;
        }

        #region ICalendarReader Members
        /// <summary>
        /// Gets the background worker used to fetch calendar entries in the background.
        /// </summary>
        /// <value></value>
        public BackgroundWorker FetchCalendarWorker { get; private set; }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            //if (_calendarEntries != null && _calendarEntries.Count > 0)
            //{
            //    worker.ReportProgress(100);
            //    e.Result = CalendarEntries;
            //    return;
            //}
            worker.ReportProgress(0);
            List<CalendarEntry> calEntries = new List<CalendarEntry>();
            SkippedEntries = new List<SkippedEntry>();
            ISession session = null;
            try
            {
                session = CreateNotesSession();
                var db = session.GetDatabase("", ProgramSettings.Instance.NotesDatabase, false);
                
                NotesView view = db.GetView("Calendar");
                
                var entries = view.AllEntries;
                List<string> completedIds = new List<string>();
                for (int row = 0; row < entries.Count; row++)
                {
                    worker.ReportProgress(100 * row / entries.Count);
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        e.Result = calEntries;
                        return;
                    }
                    var viewEntry = entries.GetNthEntry(row);
                    if (completedIds.Contains(viewEntry.NoteID))
                        continue;
                    completedIds.Add(viewEntry.NoteID);
                    SkippedEntry skippedEntry;
                    CalendarEntry calEntry = CreateCalendarEntry(viewEntry, out skippedEntry);
                    if (calEntry != null)
                        calEntries.Add(calEntry);
                    else
                    {
                        SkippedEntries.Add(skippedEntry);
                    }
                }
                e.Result = calEntries;
                CalendarEntries = calEntries;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Failed to retrieve calendar entries: " + ex.Message, ex);
            }
            finally
            {
                if (session != null)
                    Marshal.FinalReleaseComObject(session);
                worker.ReportProgress(100);
            }
        }

        /// <summary>
        /// Gets the collection of calendar entries that were skipped.
        /// </summary>
        /// <value>The skipped entries.</value>
        public ICollection<SkippedEntry> SkippedEntries { get; private set; }
        /// <summary>
        /// Gets or sets the calendar entries.
        /// </summary>
        /// <value>The calendar entries.</value>
        public IEnumerable<CalendarEntry> CalendarEntries
        {
            get { return _calendarEntries; }
            private set { _calendarEntries = (List<CalendarEntry>)value; }
        }

        /// <summary>
        /// Gets the number of calendar entries that was skipped while reading the calendar.
        /// </summary>
        /// <value>The number of skipped entries.</value>
        public int NumberOfSkippedEntries
        {
            get 
            {
                if (SkippedEntries == null)
                    return 0;
                return SkippedEntries.Count;
            }
        }

        /// <summary>
        /// Gets the version of Lotus Notes that is installed
        /// </summary>
        public string CalendarAppVersion { get; private set; }
        
        #endregion
    }
}
