using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Domino;
using System.Diagnostics;
using System.ComponentModel;
using System.Security;

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

        /// <summary>
        /// Gets or sets the database file to read calendar entries from.
        /// </summary>
        public string DatabaseFile { get; set; }
        /// <summary>
        /// Gets or sets the password required to access the <see cref="DatabaseFile"/>.
        /// </summary>
        public string Password { get; set; }
        
        private ISession CreateNotesSession()
        {
            var session = new NotesSessionClass();
            session.Initialize(Password);
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
        /// Creates the calendar entry from the provided Lotus Notes calendar entry.
        /// </summary>
        /// <param name="notesEntry">The notes entry.</param>
        /// <remarks>
        /// More details about Notes API: http://www-01.ibm.com/support/docview.wss?rs=463&context=SSKTMJ&context=SSKTWP&dc=DB520&dc=D600&dc=DB530&dc=D700&dc=DB500&dc=DB540&dc=DB510&dc=DB550&q1=1229486&uid=swg21229486&loc=en_US&cs=utf-8&lang=en
        /// </remarks>
        /// <returns></returns>
        private static CalendarEntry CreateCalendarEntry(NotesViewEntry notesEntry)
        {
            TimeSpan startTZOffset  = TimeSpan.FromTicks(0);
            TimeSpan endTZOffset    = TimeSpan.FromTicks(0);
            TimeSpan localTZOffset  = TimeSpan.FromTicks(0);
            CalendarEntry newEntry = new CalendarEntry();
            newEntry.NotesID = notesEntry.UniversalID;
            NotesDocument doc = notesEntry.Document;
            Debug.Assert(doc.UniversalID == notesEntry.UniversalID);
            Dictionary<string, string> stringItems = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
            Dictionary<string, DateTime> dateItems = new Dictionary<string, DateTime>(StringComparer.InvariantCultureIgnoreCase);
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
                    dateItems.Add(item.Name, (DateTime) item.DateTimeValue.LSGMTTime);
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
                        newEntry.Occurrences.Add((DateTime)nTime.LSGMTTime);
                    }
                }
                else if (item.Name == "StartTimeZone")
                {
                    startTZOffset = GetTimeZoneDiff(item);
                }
                else if (item.Name == "EndTimeZone")
                {
                    endTZOffset = GetTimeZoneDiff(item);
                }
                else if (item.Name == "LocalTimeZone")
                {
                    localTZOffset = GetTimeZoneDiff(item);
                }
            }
            // sanity check
            if (stringItems.ContainsKey("TaskType") && !stringItems.ContainsKey("AppointmentType"))
                // It's probably a TODO or Followup, ignore it
                return null;
            if (stringItems.ContainsKey("Body"))
                newEntry.Body = stringItems["Body"];

            if (stringItems.ContainsKey("Subject"))
                newEntry.Subject = stringItems["Subject"];
            else
                newEntry.Subject = "(no subject)";
            if (stringItems.ContainsKey("Location"))
                newEntry.Location = stringItems["Location"];
            if (stringItems.ContainsKey("Room"))
            {
                if (!String.IsNullOrEmpty(newEntry.Location))
                    newEntry.Location = String.Format("{0}, {1}", newEntry.Location, stringItems["Room"]);
                else
                    newEntry.Location = stringItems["Room"];
            }
            if (stringItems.ContainsKey("SendTo"))
                newEntry.Participants.AddRange(GetRecipentList(stringItems["SendTo"]));
            if (stringItems.ContainsKey("CopyTo"))
                newEntry.OptionalParticipants.AddRange(GetRecipentList(stringItems["CopyTo"]));
            if (stringItems.ContainsKey("From"))
                newEntry.From = stringItems["From"];

            if (dateItems.ContainsKey("StartDateTime"))
                newEntry.StartTime = dateItems["StartDateTime"];
            else if (dateItems.ContainsKey("StartDate"))
            {
                newEntry.StartTime = dateItems["StartDate"];
                if (dateItems.ContainsKey("StartTime"))
                    newEntry.UpdateStartTime(dateItems["StartTime"].TimeOfDay);
                newEntry.StartTime = newEntry.StartTime.Add(startTZOffset);
            }
            else if (dateItems.ContainsKey("StartTime"))
            {
                Debugger.Break(); // This is probably an exception
                newEntry.StartTime = dateItems["StartTime"].Add(startTZOffset);
                Debug.Assert(newEntry.StartTime.Date != default(DateTime).Date);
            }

            if (dateItems.ContainsKey("EndDateTime"))
                newEntry.EndTime = dateItems["EndDateTime"];
            else if (dateItems.ContainsKey("EndDate"))
            {                
                newEntry.EndTime = dateItems["EndDate"];
                if (dateItems.ContainsKey("EndTime"))
                    newEntry.UpdateEndTime(dateItems["EndTime"].TimeOfDay);
                
                newEntry.EndTime = newEntry.EndTime.Add(endTZOffset);
            }
            else if (dateItems.ContainsKey("EndTime"))
            {
                Debugger.Break();
                newEntry.EndTime = dateItems["EndTime"].Add(endTZOffset);
                Debug.Assert(newEntry.EndTime.Date != default(DateTime).Date);
            }
            // Notes stores timezone offset in the inverted form (e.g. CET would be -1, instead of +1)
            newEntry.SetEndTimeZoneFromOffset(TimeSpan.FromTicks(endTZOffset.Ticks * -1));
            newEntry.SetStartTimeZoneFromOffset(TimeSpan.FromTicks(startTZOffset.Ticks * -1));
            
            if (stringItems["AppointmentType"] == "2")
            {
                /* All Day Event */
                newEntry.IsAllDay = true;
                newEntry.StartTimeZone = TimeZoneInfo.Local;
                newEntry.EndTimeZone = TimeZoneInfo.Local;
                // Reset the time to zero to reflect that it is all day event
                newEntry.StartTime = TimeZoneInfo.ConvertTimeToUtc(newEntry.StartTime.Date);
                newEntry.EndTime = TimeZoneInfo.ConvertTimeToUtc(newEntry.EndTime.Date.AddDays(1));
            }
            else if (stringItems["AppointmentType"] == "1")
            {
                // Anniversary
                newEntry.StartTime = dateItems["StartTime-local"];
                newEntry.EndTime = dateItems["EndTime-local"];
                newEntry.IsAllDay = true;
            }
            if (!stringItems.ContainsKey("OrgRepeat") && newEntry.Occurrences.Count > 1)
            {
                newEntry.Occurrences.Clear();
                newEntry.Occurrences.Add(newEntry.StartTime);
            }
            //if (newEntry.Subject.Contains("test-two"))
            //    Debugger.Break();
            Debug.Assert(newEntry.Occurrences.Count > 0);
            Debug.Assert(newEntry.NotesID != null && newEntry.NotesID.Length > 0);

            return newEntry;
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
            int start = notesName.IndexOf("CN=") + "CN=".Length;
            if (start == -1)
                return notesName;
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
            NumberOfSkippedEntries = 0;
            try
            {
                var session = CreateNotesSession();
                var db = session.GetDatabase("", DatabaseFile, false);
                
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
                    CalendarEntry calEntry = CreateCalendarEntry(viewEntry);
                    if (calEntry != null)
                        calEntries.Add(calEntry);
                    else
                        NumberOfSkippedEntries++;
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
                worker.ReportProgress(100);
            }
        }

        public IEnumerable<CalendarEntry> CalendarEntries
        {
            get { return _calendarEntries; }
            private set { _calendarEntries = (List<CalendarEntry>)value; }
        }

        public int NumberOfSkippedEntries { get; private set; }
        #endregion
    }
}
