﻿using System;
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
        Application outlookApp;
        public OutlookManager()
        {
            outlookApp = new ApplicationClass();
            FetchCalendarWorker = new BackgroundWorker();
            FetchCalendarWorker.WorkerReportsProgress = true;
            FetchCalendarWorker.WorkerSupportsCancellation = true;
            FetchCalendarWorker.DoWork += new DoWorkEventHandler(FetchCalendarWorker_DoWork);

            MergeCalendarWorker = new BackgroundWorker();
            MergeCalendarWorker.WorkerReportsProgress = true;
            MergeCalendarWorker.DoWork += new DoWorkEventHandler(MergeCalendarWorker_DoWork);
        }


        private MAPIFolder GetCalendarFolder()
        {
            NameSpace outlookNS = outlookApp.GetNamespace("MAPI");
            return outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);            
        }

        private static CalendarEntry CreateCalendarEntry(AppointmentItem outlookEntry)
        {
            CalendarEntry newEntry = new CalendarEntry();
            newEntry.IsAllDay = outlookEntry.AllDayEvent;
            newEntry.Body = outlookEntry.Body;
            newEntry.Subject = outlookEntry.Subject;
            newEntry.Location = outlookEntry.Location;
            newEntry.StartTime = outlookEntry.StartUTC;
            newEntry.EndTime = outlookEntry.EndUTC;
            //newEntry.StartTimeZone = outlookEntry.StartTimeZone;
            //newEntry.EndTimeZone = outlookEntry.EndTimeZone;
            foreach (Recipient recipent in outlookEntry.Recipients)
                newEntry.Participants.Add(recipent.Name);
            if (outlookEntry.OptionalAttendees != null)
                newEntry.OptionalParticipants.Add(outlookEntry.OptionalAttendees);
            if (outlookEntry.Categories != null)
                newEntry.Categories.AddRange(outlookEntry.Categories.Split(';'));
            if (outlookEntry.IsRecurring)
            {
                var pattern = outlookEntry.GetRecurrencePattern();
                var rType = pattern.RecurrenceType;
                Debug.WriteLine(pattern.Occurrences);
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
        private void UpdateRecurrencePattern(AppointmentItem olItem, IList<DateTime> occurences)
        {
            if (occurences.Count == 1)
                return;
            var analyzer = new RepeatPatternAnalyzer(occurences);
            RecurrencePattern pattern = olItem.GetRecurrencePattern();
            if (analyzer.IsDaily)
                pattern.RecurrenceType = OlRecurrenceType.olRecursDaily;
            else if (analyzer.IsWeekly)
            {
                pattern.RecurrenceType = OlRecurrenceType.olRecursWeekly;
                switch (occurences[0].DayOfWeek)
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
            else if (analyzer.IsMonthly)
            {
                pattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                pattern.DayOfMonth = occurences[0].Day;
            }
            else if (analyzer.IsYearly)
            {
                pattern.RecurrenceType = OlRecurrenceType.olRecursYearly;
                pattern.DayOfMonth = occurences[0].Day;
                pattern.MonthOfYear = occurences[0].Month;
            }
            else
                return;
            pattern.Interval = analyzer.Interval;
            pattern.StartTime = occurences[0];
            pattern.EndTime = occurences[occurences.Count - 1];
        }
        /// <summary>
        /// Updates the specified appointmentitem with data from the provided CalendarEntry.
        /// </summary>
        /// <param name="olItem">The outlook item to update.</param>
        /// <param name="entry">The calendar entry to read updated information from.</param>
        private void UpdateEntry(AppointmentItem olItem, CalendarEntry entry)
        {
            olItem.StartTimeZone = outlookApp.TimeZones[entry.StartTimeZone.Id];
            olItem.EndTimeZone = outlookApp.TimeZones[entry.EndTimeZone.Id];
            olItem.Subject = entry.Subject;
            olItem.Body = entry.Body;
            olItem.Location = entry.Location;
            
            foreach (Recipient rcp in olItem.Recipients)
                rcp.Delete();
            foreach (var name in entry.Participants)
                olItem.Recipients.Add(name);
            olItem.OptionalAttendees = String.Join(", ", entry.OptionalParticipants.ToArray());
            olItem.Start = TimeZoneInfo.ConvertTimeFromUtc(entry.StartTime, TimeZoneInfo.Local);
            olItem.End = TimeZoneInfo.ConvertTimeFromUtc(entry.EndTime, TimeZoneInfo.Local);
            olItem.UnRead = false;
            olItem.ReminderOverrideDefault = true;
            olItem.ReminderSet = false;
            olItem.AllDayEvent = entry.IsAllDay;
            UpdateRecurrencePattern(olItem, entry.Occurrences);
            if (!entry.IsRepeating)
            {
                olItem.Save();
            }
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
                    }
                    else if (modification.Modification == Modification.New)
                    {
                        var olItem = (AppointmentItem)calendarFolder.Items.Add(OlItemType.olAppointmentItem);
                        UpdateEntry(olItem, modification.Entry);
                        modification.Entry.OutlookID = olItem.GlobalAppointmentID;
                    }
                    else if (modification.Modification == Modification.Removed)
                    {
                        var olItem = GetExistingAppointmentItem(modification.Entry.OutlookID, calendarFolder);
                        olItem.Delete();
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

        #region ICalendarReader Members
        /// <summary>
        /// Gets the background worker used to fetch calendar entries in the background.
        /// </summary>
        /// <value></value>
        public BackgroundWorker FetchCalendarWorker { get; private set; }

        void FetchCalendarWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            worker.ReportProgress(0);
            List<CalendarEntry> calEntries = new List<CalendarEntry>();
            try
            {
                MAPIFolder calendarFolder = GetCalendarFolder();
                int i = 0;
                foreach (AppointmentItem item in calendarFolder.Items)
                {
                    var calEntry = CreateCalendarEntry(item);
                    if (calEntry.OutlookID != null)
                        // TODO: report error when that mechanism exists
                        calEntries.Add(calEntry);
                    i++;
                    if (worker.CancellationPending)
                    {
                        e.Result = calEntries;
                        return;
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

        #endregion
    }
}