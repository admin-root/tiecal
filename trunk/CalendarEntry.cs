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
using Microsoft.Office.Interop.Outlook;

namespace TieCal
{    
    /// <summary>
    /// Application neutral class that represents a calendar entry.
    /// </summary>
    public class CalendarEntry
    {
        private DateTime _endTime;
        private DateTime _startTime;

        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarEntry"/> class.
        /// </summary>
        public CalendarEntry()
        {
            Participants = new List<string>();
            Occurrences = new List<DateTime>();
            Categories = new List<string>();
            OptionalParticipants = new List<string>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarEntry"/> class. 
        /// </summary>
        /// <param name="copy">The calendar entry to copy information from.</param>
        /// <param name="newStartTime">The new start time to use for the created calendar entry.</param>
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
        /// <summary>
        /// Sets the <see cref="StartTimeZone"/> from a UTC offset.
        /// </summary>
        public void SetStartTimeZoneFromOffset(TimeSpan offset)
        {
            StartTimeZone = GetTimeZoneInfoFromOffset(offset);
        }

        /// <summary>
        /// Sets the <see cref="EndTimeZone"/> from a UTC offset.
        /// </summary>
        public void SetEndTimeZoneFromOffset(TimeSpan offset)
        {
            EndTimeZone = GetTimeZoneInfoFromOffset(offset);
        }

        /// <summary>
        /// Gets or sets the person this entry originates from.
        /// </summary>
        public string From { get; set; }
        /// <summary>
        /// Gets or sets the list of categories this entry belongs to.
        /// </summary>
        public List<string> Categories { get; set; }
        /// <summary>
        /// Gets or sets the ID this entry has in Outlook. This property is only set after entries has been read from Outlook or this entry has been successfully merged with Outlook.
        /// </summary>
        public string OutlookID { get; set; }
        /// <summary>
        /// Gets or sets the ID this entry has in Lotus Notes. This property is only set for items read from the <see cref="NotesReader"/> class.
        /// </summary>
        public string NotesID { get; set; }
        /// <summary>
        /// Gets or sets the subject/title of the calendar entry.
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// Gets or sets the full description of the calendar entry.
        /// </summary>
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
        /// <summary>
        /// Gets or sets a value indicating whether this calendar entry is an all day event.
        /// </summary>
        public bool IsAllDay { get; set; }
        /// <summary>
        /// Gets or sets the list of dates when this entry occurrs. For non-repeating events, this contains only one item (the <see cref="StartTime"/>)
        /// </summary>
        public List<DateTime> Occurrences { get; set; }
        /// <summary>
        /// Gets the duration of the meeting.
        /// </summary>
        public TimeSpan Duration { get { return EndTime - StartTime; } }
        /// <summary>
        /// Gets a value indicating whether this calendar entry is a repeating event. To get the repeating events, either look at the <see cref="Occurrences"/> property or use the <see cref="RepeatPatternAnalyzer"/> to get a proper repeat pattern.
        /// </summary>
        public bool IsRepeating { get { return Occurrences.Count > 1; } }
        /// <summary>
        /// Gets or sets the time zone that the <see cref="StartTime"/> is expressed in
        /// </summary>
        public TimeZoneInfo StartTimeZone { get; set; }
        /// <summary>
        /// Gets or sets the time zone that the <see cref="EndTime"/> is expressed in.
        /// </summary>
        public TimeZoneInfo EndTimeZone { get; set; }
        
        /// <summary>
        /// Gets or sets the start time. This should always be in UTC time. Use the <see cref="StartTimeLocal"/> property if you want the time in the local timezone
        /// </summary>
        /// <value>The start time, in UTC.</value>
        public DateTime StartTime { get { return _startTime; } set { _startTime = value; } }
        /// <summary>
        /// Gets or sets the end time. This should always be in UTC time. Use the <see cref="EndTimeLocal"/> property if you want the time in the local timezone
        /// </summary>
        /// <value>The end time, in UTC.</value>
        public DateTime EndTime { get { return _endTime; } set { _endTime = value; } }

        /// <summary>
        /// Gets the start time expressed in the current (local) timezone.
        /// </summary>
        public DateTime StartTimeLocal
        {
            get
            {
                return TimeZoneInfo.ConvertTimeFromUtc(StartTime, TimeZoneInfo.Local);
            }
        }
        /// <summary>
        /// Gets the end time expressed in the current (local) timezone.
        /// </summary>
        public DateTime EndTimeLocal
        {
            get
            {
                return TimeZoneInfo.ConvertTimeFromUtc(EndTime, TimeZoneInfo.Local);
            }
        }
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

        /// <summary>
        /// Gets the list of properties that are different between this calendar entry and the <paramref name="other"/>.
        /// </summary>
        /// <param name="other">The calendar entry to compare against.</param>
        /// <returns>A list of property names that differ</returns>
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

        /// <summary>
        /// Gets a string representation of this calendar entry.
        /// </summary>        
        public override string ToString()
        {
            return Subject;
        }
    }

    /// <summary>
    /// Analyzer for a list of <see cref="DateTime"/> objects which tries to find a pattern that can be applied when creating/updating outlook entries
    /// </summary>
    public class RepeatPatternAnalyzer
    {
        private IList<DateTime> Occurrences { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="RepeatPatternAnalyzer"/> class by analyizing the list of occurrences.
        /// </summary>
        /// <param name="occurrences">The occurrences to analyze in order to calculate the repeat pattern.</param>
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
        private bool SameMinute { get; set; }
        private bool SameHour { get; set; }
        private bool SameTime { get { return SameMinute && SameHour; } }
        private bool SameDayOfMonth { get; set; }
        private bool SameDayOfWeek { get; set; }
        private bool SameMonth { get; set; }

        /// <summary>
        /// Gets a value indicating whether this is a yearly event.
        /// </summary>
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

        /// <summary>
        /// Gets a value indicating whether this is a monthly event.
        /// </summary>
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
        /// <summary>
        /// Gets a value indicating whether this is a weekly event.
        /// </summary>
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

        /// <summary>
        /// Gets a value indicating whether this is a daily event.
        /// </summary>
        public bool IsDaily
        {
            get
            {
                if (!SameTime)
                    return false;
                if (SameDayOfWeek || SameDayOfMonth)
                    return false;
                
                return true;
            }
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current repeat pattern.
        /// </summary>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            if (!IsValid)
                return "Invalid Repeat Pattern";
            if (IsDaily)
                sb.Append("Daily event.");
            if (IsWeekly)
                sb.Append("Weekly event.");
            if (IsMonthly)
                sb.Append("Monthly event.");
            if (IsYearly)
                sb.Append("Yearly event.");
            if (sb.Length == 0)
                sb.Append("Unknown repeat pattern");
            sb.AppendFormat("{0} occurrences, interval: {1}", NumRepeats, Interval);

            return sb.ToString();
        }
    }
}
