using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Domino;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;

namespace TieCal
{    
    public class CalendarEntry
    {
        public CalendarEntry()
        {
            Participants = new List<string>();
            Occurrences = new List<DateTime>();
            Categories = new List<string>();
            OptionalParticipants = new List<string>();
        }

        public CalendarEntry(CalendarEntry copy, DateTime newStartTime)
            : this()
        {
            Body = copy.Body;
            Subject = copy.Subject;
            NotesID = copy.NotesID;
            Occurrences.Add(newStartTime);
            StartTime = newStartTime;
            EndTime = newStartTime + copy.Duration;
            Participants.AddRange(copy.Participants);
            OptionalParticipants.AddRange(copy.OptionalParticipants);
            Categories.AddRange(copy.Categories);
            OutlookID = copy.OutlookID;
            IsAllDay = copy.IsAllDay;
            Location = copy.Location;

        }

        /// <summary>
        /// Updates the start date with the specified date information while maintaining the timestamp.
        /// </summary>
        /// <param name="datePart">The date part.</param>
        public void UpdateStartDate(DateTime datePart)
        {
            StartTime = new DateTime(datePart.Year, datePart.Month, datePart.Day, StartTime.Hour, StartTime.Minute, StartTime.Second); 
        }

        /// <summary>
        /// Updates the provided datetime with the specified time information while maintaining the original date.
        /// </summary>
        /// <param name="timePart">The time part.</param>
        public void UpdateStartTime(TimeSpan timePart)
        {
            StartTime = new DateTime(StartTime.Year, StartTime.Month, StartTime.Day, timePart.Hours, timePart.Minutes, timePart.Seconds);            
        }

        /// <summary>
        /// Updates the start date with the specified date information while maintaining the timestamp.
        /// </summary>
        /// <param name="datePart">The date part.</param>
        public void UpdateEndDate(DateTime datePart)
        {
            EndTime = new DateTime(datePart.Year, datePart.Month, datePart.Day, EndTime.Hour, EndTime.Minute, EndTime.Second);
        }

        /// <summary>
        /// Updates the provided datetime with the specified time information while maintaining the original date.
        /// </summary>
        /// <param name="timePart">The time part.</param>
        public void UpdateEndTime(TimeSpan timePart)
        {
            EndTime = new DateTime(EndTime.Year, EndTime.Month, EndTime.Day, timePart.Hours, timePart.Minutes, timePart.Seconds);
        }

        /// <summary>
        /// Gets the time zone info from offset.
        /// </summary>
        /// <param name="offset">The offset.</param>
        /// <returns></returns>
        private TimeZoneInfo GetTimeZoneInfoFromOffset(TimeSpan offset)
        {            
            foreach (var zone in TimeZoneInfo.GetSystemTimeZones())
                if (zone.BaseUtcOffset == offset)
                    return zone;
            return TimeZoneInfo.Utc;
        }
        public void SetStartTimeZoneFromOffset(TimeSpan offset)
        {
            StartTimeZone = GetTimeZoneInfoFromOffset(offset);
        }

        public void SetEndTimeZoneFromOffset(TimeSpan offset)
        {
            EndTimeZone = GetTimeZoneInfoFromOffset(offset);
        }

        public IEnumerable<CalendarEntry> GetRepeatingEntries()
        {
            List<CalendarEntry> lst = new List<CalendarEntry>();
            if (!this.IsRepeating)
                return lst;
            foreach (DateTime repeatTime in Occurrences)
                lst.Add(new CalendarEntry(this, repeatTime));
            return lst;
        }

        public string From { get; set; }
        public List<string> Categories { get; set; }
        public string OutlookID { get; set; }
        public string NotesID { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        /// <summary>
        /// Gets or sets the location where the meeting is.
        /// </summary>
        public string Location { get; set; }
        /// <summary>
        /// Gets or sets the list of participants.
        /// </summary>
        public List<string> Participants { get; set; }
        /// <summary>
        /// Gets or sets the list of optional participants.
        /// </summary>
        public List<string> OptionalParticipants { get; set; }
        public bool IsAllDay { get; set; }
        public List<DateTime> Occurrences { get; set; }
        /// <summary>
        /// Gets the duration of the meeting.
        /// </summary>
        public TimeSpan Duration { get { return EndTime - StartTime; } }
        public bool IsRepeating { get { return Occurrences.Count > 1; } }
        private DateTime _startTime;
        /// <summary>
        /// Gets or sets the time zone that the <see cref="StartTime"/> is expressed in
        /// </summary>
        /// <value>The start time zone.</value>
        public TimeZoneInfo StartTimeZone { get; set; }
        /// <summary>
        /// Gets or sets the time zone that the <see cref="EndTime"/> is expressed in.
        /// </summary>
        /// <value>The end time zone.</value>
        public TimeZoneInfo EndTimeZone { get; set; }
        
        /// <summary>
        /// Gets or sets the start time. This should always be in UTC time
        /// </summary>
        /// <value>The start time, in UTC.</value>
        public DateTime StartTime { get { return _startTime; } set { _startTime = value; } }
        private DateTime _endTime;
        /// <summary>
        /// Gets or sets the end time. This should always be in UTC time
        /// </summary>
        /// <value>The end time, in UTC.</value>
        public DateTime EndTime { get { return _endTime; } set { _endTime = value; } }

        /// <summary>
        /// Gets a value indicating whether this calendar entry occurs in the specified interval.
        /// </summary>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <returns></returns>
        public bool OccursInInterval(DateTime start, DateTime end)
        {
            foreach (var occurrence in Occurrences)
                if (occurrence > start && occurrence < end)
                    return true;
            return false;
        }

        /// <summary>
        /// Determines whether the this calendar entry is equivalents to the specified calendar entry for the sake of merging.
        /// </summary>
        /// <param name="other">The calendar entry to compare against.</param>
        /// <returns>True if they are equivalent (no merging required), false otherwise</returns>
        /// <seealso cref="DiffersFrom"/>
        public bool EquivalentTo(CalendarEntry other)
        {
            if (Subject != other.Subject)
                return false;
            if (Body != other.Body)
                return false;
            if (StartTime != other.StartTime || EndTime != other.EndTime)
                return false;
            if (Location != other.Location)
                return false;
            return true;
        }

        /// <summary>
        /// Determines whether this calendar entry differs from the specified calendar entry for the sake of merging.
        /// </summary>
        /// <param name="other">The calendar entry to compare against.</param>
        /// <returns>True if they are different (merging is required), false otherwise</returns>
        /// <seealso cref="EquivalentTo"/>
        public bool DiffersFrom(CalendarEntry other)
        {
            return !EquivalentTo(other);
        }

        public List<string> GetDifferences(CalendarEntry other)
        {
            List<string> diffs = new List<string>();
            if (this.EquivalentTo(other))
                return diffs;
            if (Subject != other.Subject)
                diffs.Add("Subject");
            if (Body != other.Body)
                diffs.Add("Body");
            if (StartTime != other.StartTime)
                diffs.Add("StartTime");
            if (EndTime != other.EndTime)
                diffs.Add("EndTime");
            if (Location != other.Location)
                diffs.Add("Location");
            return diffs;
        }
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Subject: " + Subject);
            sb.AppendLine("From: " + From);
            sb.Append("Categories: ");
            foreach (string cat in Categories)
                sb.Append(cat + ", ");
            sb.AppendLine();
            sb.AppendLine("Location: " + Location);
            sb.AppendLine("Start: " + StartTime);
            sb.AppendLine("End:   " + EndTime);
            sb.AppendLine("Notes ID: " + NotesID);
            sb.AppendLine("Outlook ID: " + OutlookID);
            sb.AppendLine("Num Participants: " + Participants.Count);
            sb.AppendLine("Num Opt. Participants: " + OptionalParticipants.Count);
            sb.AppendLine("Repeating: " + IsRepeating);
            if (IsRepeating)
                sb.AppendLine(" Number of repeats: " + Occurrences.Count);
            
            return sb.ToString();
        }
    }

    /// <summary>
    /// Analyzer for a list of <see cref="DateTime"/> objects which tries to find a pattern that can be applied when creating/updating outlook entries
    /// </summary>
    public class RepeatPatternAnalyzer
    {
        private IList<DateTime> Occurrences { get; set; }
        public RepeatPatternAnalyzer(IList<DateTime> occurrences)
        {
            if (occurrences.Count < 2)
                throw new ArgumentException("Cannot calculate repeat pattern unless there's at least two occurrences", "occurrences");
            Occurrences = occurrences;
            SameMinute = true;
            SameHour = true;
            SameDayOfMonth = true;
            SameDayOfWeek = true;
            SameMonth = true;

            for (int i = 1; i < occurrences.Count; i++)
            {
                var prev = occurrences[i - 1];
                var cur = occurrences[i];
                if (cur.Minute != prev.Minute)
                    SameMinute = false;
                if (cur.Hour != prev.Hour)
                    SameHour = false;
                if (cur.DayOfWeek != prev.DayOfWeek)
                    SameDayOfWeek = false;
                if (cur.Day != prev.Day)
                    SameDayOfMonth = false;
                if (cur.Month != prev.Month)
                    SameMonth = false;
            }
            
            if (!IsValid)
                throw new ArgumentException ("The occurrences does not map to a valid repeat pattern", "occurrences");
        }

        /// <summary>
        /// Gets the interval between repeats. The unit depends on the IsDaily, IsWeekly etc. properties
        /// </summary>
        public int Interval { get { return 1; } }

        /// <summary>
        /// Gets a value indicating whether this instance is valid, i.e. the occurrences were mapped to something that makes sense.
        /// </summary>
        private bool IsValid
        {
            get
            {
                if (!(SameMinute || SameHour || SameDayOfMonth || SameDayOfWeek || SameMonth))
                    // No pattern was found
                    return false;
                if (!IsDaily && !IsWeekly && !IsMonthly && !IsYearly)
                    return false;
                // TODO: Maybe also make sure it's only one of the above
                return true;
            }
        }
        public int NumRepeats { get { return Occurrences.Count; } }
        public bool SameMinute { get; private set; }
        public bool SameHour { get; private set; }
        public bool SameTime { get { return SameMinute && SameHour; } }
        public bool SameDayOfMonth { get; private set; }
        public bool SameDayOfWeek { get; private set; }
        public bool SameMonth { get; private set; }

        public bool IsYearly
        {
            get
            {
                if (!SameTime)
                    return false;
                if (!SameMonth && !SameDayOfMonth)
                    return false;
                
                return true;
            }
        }

        public bool IsMonthly
        {
            get
            {
                if (!SameTime)
                    return false;
                if (!SameDayOfMonth)
                    return false;
                if (SameMonth)
                    return false;
                return true;
            }
        }
        public bool IsWeekly
        {
            get
            {
                if (!SameTime)
                    return false;
                if (!SameDayOfWeek)
                    return false;
                return true;
            }
        }

        public bool IsDaily
        {
            get
            {
                if (!SameTime)
                    return false;
                if (SameDayOfWeek || SameDayOfMonth)
                    return false;
                if (SameMonth)
                    return false;

                return true;
            }

        }
    }
}
