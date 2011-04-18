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
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.ComponentModel;

namespace TieCal
{
    /// <summary>
    /// Handles communication with Outlook
    /// </summary>
    public class OutlookManager : ICalendarReader
    {
        /// <summary>
        /// Do not use directly, use the <see cref="OutlookApp"/> property instead which is guaranteed to be non-null.
        /// </summary>
        Application outlookApp;
        public OutlookManager()
        {
            FetchCalendarWorker = new BackgroundWorker();
            FetchCalendarWorker.WorkerReportsProgress = true;
            FetchCalendarWorker.WorkerSupportsCancellation = true;
            FetchCalendarWorker.DoWork += new DoWorkEventHandler(FetchCalendarWorker_DoWork);

            MergeCalendarWorker = new BackgroundWorker();
            MergeCalendarWorker.WorkerReportsProgress = true;
            MergeCalendarWorker.DoWork += new DoWorkEventHandler(MergeCalendarWorker_DoWork);
        }

        private Application OutlookApp
        {
            get
            {
                if (outlookApp == null)
                    outlookApp = new ApplicationClass();
                return outlookApp;
            }
        }
        private MAPIFolder GetCalendarFolder()
        {
            NameSpace outlookNS = OutlookApp.GetNamespace("MAPI");
            return outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);            
        }

        private static CalendarEntry CreateCalendarEntry(AppointmentItem outlookEntry)
        {
            CalendarEntry newEntry = new CalendarEntry();
            newEntry.IsAllDay = outlookEntry.AllDayEvent;
            newEntry.Body = outlookEntry.Body;
            newEntry.Subject = outlookEntry.Subject;
            newEntry.Location = outlookEntry.Location;
            if (newEntry.IsAllDay)
            {
                // Timezones and exact time isn't used for all day events
                newEntry.StartTime = outlookEntry.Start;
                newEntry.EndTime = outlookEntry.End;
            }
            else
            {
                newEntry.StartTime = outlookEntry.StartUTC;
                newEntry.EndTime = outlookEntry.EndUTC;
            }
            newEntry.StartTimeZone = TimeZoneInfo.Local; //outlookEntry.StartTimeZone;
            newEntry.EndTimeZone = TimeZoneInfo.Local;
            foreach (Recipient recipent in outlookEntry.Recipients)
                newEntry.Participants.Add(recipent.Name);
            if (outlookEntry.OptionalAttendees != null)
                newEntry.OptionalParticipants.Add(outlookEntry.OptionalAttendees);
            if (outlookEntry.Categories != null)
                newEntry.Categories.AddRange(outlookEntry.Categories.Split(';'));
            if (outlookEntry.IsRecurring)
            {
                var pattern = outlookEntry.GetRecurrencePattern();
                newEntry.SetRepeatPattern(pattern);
            }
            newEntry.OutlookID = outlookEntry.GlobalAppointmentID;
            return newEntry;
        }

        private AppointmentItem GetExistingAppointmentItem(string outlookID, MAPIFolder calendarFolder)
        {
            if (outlookID == null)
                throw new ArgumentNullException("outlookID");
            if (calendarFolder == null)
                throw new ArgumentNullException("calendarFolder");
            foreach (AppointmentItem item in calendarFolder.Items)
                if (item.GlobalAppointmentID == outlookID)
                    return item;
            throw new ApplicationException("No outlook appointment with id=" + outlookID + " exists");
        }
        private AppointmentItem GetOrCreateAppointmentItem(string outlookID, MAPIFolder calendarFolder)
        {
            if (outlookID != null)
            {
                foreach (AppointmentItem item in calendarFolder.Items)
                    if (item.GlobalAppointmentID == outlookID)
                        return item;
            }

            return (AppointmentItem)calendarFolder.Items.Add(OlItemType.olAppointmentItem);
        }

        private void UpdateRecurrencePattern(AppointmentItem olItem, CalendarEntry entry)
        {            
            //if (olItem.Subject.Contains("Wasabi"))
            //    Debugger.Break();
            RecurrencePattern pattern = olItem.GetRecurrencePattern();
            if (entry.RepeatPattern.IsDaily)
                pattern.RecurrenceType = OlRecurrenceType.olRecursDaily;
            else if (entry.RepeatPattern.IsWeekly)
            {
                pattern.RecurrenceType = OlRecurrenceType.olRecursWeekly;
                switch (entry.RepeatPattern.DayOfWeek)
                {
                    case DayOfWeek.Friday:
                        pattern.DayOfWeekMask = OlDaysOfWeek.olFriday;
                        break;
                    case DayOfWeek.Monday:
                        pattern.DayOfWeekMask = OlDaysOfWeek.olMonday;
                        break;
                    case DayOfWeek.Saturday:
                        pattern.DayOfWeekMask = OlDaysOfWeek.olSaturday;
                        break;
                    case DayOfWeek.Sunday:
                        pattern.DayOfWeekMask = OlDaysOfWeek.olSunday;
                        break;
                    case DayOfWeek.Thursday:
                        pattern.DayOfWeekMask = OlDaysOfWeek.olThursday;
                        break;
                    case DayOfWeek.Tuesday:
                        pattern.DayOfWeekMask = OlDaysOfWeek.olTuesday;
                        break;
                    case DayOfWeek.Wednesday:
                        pattern.DayOfWeekMask = OlDaysOfWeek.olWednesday;
                        break;
                    default:
                        Debugger.Break();
                        break;
                }
            }
            else if (entry.RepeatPattern.IsWeeklyMultipleDays)
            {
                pattern.RecurrenceType = OlRecurrenceType.olRecursWeekly;
                // Typecasting is ok here, because the value of both enums match
                pattern.DayOfWeekMask = (OlDaysOfWeek) entry.RepeatPattern.DayOfWeekMask;
            }
            else if (entry.RepeatPattern.IsMonthly)
            {
                pattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                pattern.DayOfMonth = entry.RepeatPattern.DayOfMonth;
            }
            else if (entry.RepeatPattern.IsYearly)
            {
                pattern.RecurrenceType = OlRecurrenceType.olRecursYearly;
                pattern.DayOfMonth = entry.RepeatPattern.DayOfMonth;
                pattern.MonthOfYear = entry.RepeatPattern.MonthOfYear;
            }
            else
                return;
            pattern.Interval = entry.RepeatPattern.Interval;
            if (entry.IsAllDay)
                pattern.PatternStartDate = entry.RepeatPattern.FirstOccurrence.Date;
            else
            {
                pattern.PatternStartDate = entry.RepeatPattern.FirstOccurrence;
                pattern.StartTime = entry.RepeatPattern.FirstOccurrence;
            }
            pattern.Occurrences = entry.RepeatPattern.NumRepeats;
        }
        
        /// <summary>
        /// Updates the specified appointmentitem with data from the provided CalendarEntry.
        /// </summary>
        /// <param name="olItem">The outlook item to update.</param>
        /// <param name="entry">The calendar entry to read updated information from.</param>
        private void UpdateEntry(AppointmentItem olItem, CalendarEntry entry)
        {
            olItem.StartTimeZone = OutlookApp.TimeZones[entry.StartTimeZone.Id];
            olItem.EndTimeZone = OutlookApp.TimeZones[entry.EndTimeZone.Id];
            olItem.Subject = entry.Subject;
            olItem.Body = entry.Body;
            olItem.Location = entry.Location;
            // Add participants to outlook. I don't dare use this code since outlook may start sending them emails in case they exist in the address book
            // .. and also, iTunes ignores this info anyway so there's really no use syncing it
            //foreach (Recipient rcp in olItem.Recipients)
            //    rcp.Delete();
            //foreach (var name in entry.Participants)
            //    olItem.Recipients.Add(name);
            //olItem.OptionalAttendees = String.Join(", ", entry.OptionalParticipants.ToArray());
            if (olItem.IsRecurring && !entry.IsRepeating)
            {
                // Entry in notes is no longer a repeating event
                olItem.ClearRecurrencePattern();
            }
            if (!olItem.IsRecurring || olItem.GlobalAppointmentID == null)
            {
                // We're not allowed to modify these properties of existing repeating events (it's ok for new ones though)
                olItem.Start = entry.StartTimeLocal;
                olItem.End = entry.EndTimeLocal;
                olItem.AllDayEvent = entry.IsAllDay;
            }

            olItem.UnRead = false;
            if (entry.StartTimeLocal < DateTime.Now || entry.IsAllDay ||
                ProgramSettings.Instance.ReminderMode == ReminderMode.NoReminder)
            {
                olItem.ReminderOverrideDefault = true;
                olItem.ReminderSet = false;
            }
            else if (ProgramSettings.Instance.ReminderMode == ReminderMode.Custom)
            {
                olItem.ReminderOverrideDefault = true;
                olItem.ReminderMinutesBeforeStart = ProgramSettings.Instance.ReminderMinutesBeforeStart;
            }
            if (entry.IsRepeating)
                UpdateRecurrencePattern(olItem, entry);
            if (ProgramSettings.Instance.SyncRepeatingEvents == false && entry.IsRepeating == true)
                Debugger.Break();
            olItem.Save();
        }
        
        public void DeleteAllEntries()
        {
            var calendarFolder = GetCalendarFolder();

            foreach (AppointmentItem item in calendarFolder.Items)
            {
                item.Delete();
            }
        }

        void MergeCalendarWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            MergeCalendarWorker.ReportProgress(0);
            IEnumerable<ModifiedEntry> changedEntries = (IEnumerable<ModifiedEntry>) e.Argument;
            var calendarFolder = GetCalendarFolder();
            int count = changedEntries.Count();
            int num = 0;
            NumberOfMergedEntries = 0;
            foreach (var modification in changedEntries)
            {
                MergeCalendarWorker.ReportProgress((num++ * 100) / count);
                if (modification.ApplyModification == false)
                    continue;
                try
                {
                    if (modification.Modification == Modification.Modified)
                    {
                        AppointmentItem olItem = GetExistingAppointmentItem(modification.Entry.OutlookID, calendarFolder);
                        UpdateEntry(olItem, modification.Entry);
                        NumberOfMergedEntries++;
                    }
                    else if (modification.Modification == Modification.New)
                    {
                        var olItem = (AppointmentItem)calendarFolder.Items.Add(OlItemType.olAppointmentItem);
                        UpdateEntry(olItem, modification.Entry);
                        modification.Entry.OutlookID = olItem.GlobalAppointmentID;
                        NumberOfMergedEntries++;
                    }
                    else if (modification.Modification == Modification.Removed)
                    {
                        var olItem = GetExistingAppointmentItem(modification.Entry.OutlookID, calendarFolder);
                        olItem.Delete();
                        NumberOfMergedEntries++;
                    }
                    else
                    {
                        Debugger.Break();
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("Failed to merge " + modification.Modification + " entry (" + modification.Entry.Subject + "): " + ex.Message);                    
                }
            }
        }

        public BackgroundWorker MergeCalendarWorker { get; private set; }

        public int NumberOfMergedEntries { get; private set; }
        #region ICalendarReader Members
        /// <summary>
        /// Gets the background worker used to fetch calendar entries in the background.
        /// </summary>
        /// <value></value>
        public BackgroundWorker FetchCalendarWorker { get; private set; }

        void FetchCalendarWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            SkippedEntries = new List<SkippedEntry>();
            worker.ReportProgress(0);
            List<CalendarEntry> calEntries = new List<CalendarEntry>();
            try
            {
                // TODO: This can throw COM exception (HRESULT=0x8001010A, RPC_E_SERVERCALL_RETRYLATER) - we should handle that by retrying at least once...
                MAPIFolder calendarFolder = GetCalendarFolder();
                int i = 0;
                foreach (AppointmentItem item in calendarFolder.Items)
                {
                    try
                    {
                        var calEntry = CreateCalendarEntry(item);
                        if (calEntry.OutlookID == null)
                            SkippedEntries.Add(new SkippedEntry(calEntry, "No valid ID found on calendar entry"));
                        else if (calEntry.Categories != null && calEntry.Categories.Contains("nosync"))
                            SkippedEntries.Add(new SkippedEntry(calEntry, "Calendar entry was in the 'nosync' category"));
                        else
                            calEntries.Add(calEntry);
                        i++;
                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                            e.Result = calEntries;
                            return;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        // By doing catch-all here, we at least let the user sync the entries TieCal understands..
                        // TODO: proper error reporting
                        SkippedEntries.Add(new SkippedEntry(new CalendarEntry() { Subject = "(no subject)" }, "Exception while reading outlook: " + ex.Message));

                    }
                    worker.ReportProgress(100 * i / calendarFolder.Items.Count);

                }
                e.Result = calEntries;
                CalendarEntries = calEntries;
            }
            catch (System.Exception ex)
            {
                throw new ApplicationException("Failed to retrieve Outlook calendar items: " + ex.Message, ex);
            }
            finally
            {
                worker.ReportProgress(100);
            }
        }

        public IEnumerable<CalendarEntry> CalendarEntries
        {
            get;
            private set;
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
        /// Gets the collection of calendar entries that were skipped.
        /// </summary>
        /// <value>The skipped entries.</value>
        public ICollection<SkippedEntry> SkippedEntries { get; private set; }
        /// <summary>
        /// Gets the Microsoft Outlook version.
        /// </summary>
        /// <value>The version reported from the calendar application.</value>
        public string CalendarAppVersion { get { return OutlookApp.Version; } }
        #endregion
    }
}
