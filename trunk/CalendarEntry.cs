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
}
