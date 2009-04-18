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
        Application outlookApp;
        public OutlookManager()
        {
            outlookApp = new ApplicationClass();
            FetchCalendarWorker = new BackgroundWorker();
            FetchCalendarWorker.WorkerReportsProgress = true;
            FetchCalendarWorker.WorkerSupportsCancellation = false;
            FetchCalendarWorker.DoWork += new DoWorkEventHandler(worker_DoWork);
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

        private AppointmentItem GetAppointmentItem(string outlookID, MAPIFolder calendarFolder)
        {
            if (outlookID != null)
            {
                foreach (AppointmentItem item in calendarFolder.Items)
                    if (item.GlobalAppointmentID == outlookID)
                        return item;
            }
            return (AppointmentItem)calendarFolder.Items.Add(OlItemType.olAppointmentItem);
        }

        /// <summary>
        /// Merges the specified calendar entries with items already in Outlook. Each entry
        /// must have a valid OutlookID if it should be updated, otherwise a new entry will be created.
        /// </summary>
        /// <remarks>No entries in outlook will be removed in this method</remarks>
        /// <param name="entries">The entries to merge.</param>
        public void MergeCalendarEntries(IEnumerable<CalendarEntry> entries)
        {
            var calendarFolder = GetCalendarFolder();
            foreach (var entry in entries)
            {
                try
                {
                    AppointmentItem olItem = GetAppointmentItem(entry.OutlookID, calendarFolder);
                    olItem.StartTimeZone = outlookApp.TimeZones[entry.StartTimeZone.Id];
                    olItem.EndTimeZone = outlookApp.TimeZones[entry.EndTimeZone.Id];
                    olItem.Subject = entry.Subject;
                    olItem.Body = entry.Body;
                    olItem.Location = entry.Location;
                    
                    foreach (var name in entry.Participants)
                        olItem.Recipients.Add(name);
                    olItem.OptionalAttendees = String.Join(", ", entry.OptionalParticipants.ToArray());
                    olItem.Start = TimeZoneInfo.ConvertTimeFromUtc(entry.StartTime, TimeZoneInfo.Local);
                    olItem.End = TimeZoneInfo.ConvertTimeFromUtc(entry.EndTime, TimeZoneInfo.Local);
                    olItem.UnRead = false;
                    olItem.ReminderOverrideDefault = true;
                    olItem.ReminderSet = false;
                    olItem.Save();
                    entry.OutlookID = olItem.GlobalAppointmentID;
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine("Failed to handle entry: " + entry + ex.Message + Environment.NewLine + "----------------");
                }
            }
        }

        public void DeleteAllEntries()
        {
            var calendarFolder = GetCalendarFolder();

            foreach (AppointmentItem item in calendarFolder.Items)
            {
                item.Delete();
                //item.Save();
            }
        }

        #region ICalendarReader Members
        /// <summary>
        /// Gets the background worker used to fetch calendar entries in the background.
        /// </summary>
        /// <value></value>
        public BackgroundWorker FetchCalendarWorker { get; private set; }

        public void BeginFetchCalendarEntries()
        {
            FetchCalendarWorker.RunWorkerAsync();
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
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
