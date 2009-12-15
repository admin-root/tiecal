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
        /// Gets the duration of the meeting.
        /// </summary>
        public TimeSpan Duration { get { return EndTime - StartTime; } }
        /// <summary>
        /// Gets a value indicating whether this calendar entry is a repeating event. To get the repeating events, either look at the <see cref="SetRepeatingPattern"/> method.
        /// </summary>
        public bool IsRepeating { get; private set; }
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
                if (IsAllDay)
                    return StartTime;
                else
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
                if (IsAllDay)
                    return EndTime;
                else
                    return TimeZoneInfo.ConvertTimeFromUtc(EndTime, TimeZoneInfo.Local);
            }
        }

        private RepeatPattern _repeatPattern;
        public RepeatPattern RepeatPattern
        {
            get
            {
                if (!IsRepeating)
                    throw new InvalidOperationException("Cannot get repeat pattern for non-repeating events");
                return _repeatPattern;
            }
        }

        public void SetRepeatPattern(IList<DateTime> occurrences)
        {
            _repeatPattern = RepeatPattern.CreateFromOccurrences(occurrences);
            IsRepeating = true;
        }

        public void SetRepeatPattern(RecurrencePattern outlookRecurrencePattern)
        {
            _repeatPattern = RepeatPattern.CreateFromOutlookPattern(outlookRecurrencePattern);
            IsRepeating = true;
        }
        /// <summary>
        /// Gets a value indicating whether this calendar entry occurs in the specified interval.
        /// </summary>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <returns></returns>
        public bool OccursInInterval(DateTime start, DateTime end)
        {
            throw new NotImplementedException("This isn't implemented yet");
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
            if (IsAllDay != other.IsAllDay)
                return false;
            if (IsAllDay)
            {
                // Repeating events cannot be modified in outlook it seems. And also, Outlook sets StartTime to the first in the series
                // while notes sets it to the current instance.. (TODO: fix!)
                if (!IsRepeating)
                {
                    if (StartTime.Date != other.StartTime.Date ||
                        EndTime.Date != other.EndTime.Date)
                        return false;
                }
            }
            else
            {
                if (StartTime != other.StartTime || EndTime != other.EndTime)
                    return false;
            }
            if (Location != other.Location)
                return false;
            if (IsRepeating != other.IsRepeating)
                return false;
            if (IsRepeating && other.IsRepeating)
            {
                if (!RepeatPattern.EquivalentTo(other.RepeatPattern))
                    return false;
            }
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
            if (IsAllDay != other.IsAllDay)
                diffs.Add("Is All Day");
            if (IsAllDay)
            {
                if (!IsRepeating)
                {
                    if (StartTime.Date != other.StartTime.Date)
                        diffs.Add("Start Date");
                    if (EndTime.Date != other.EndTime.Date)
                        diffs.Add("End Date");
                }
            }
            else
            {
                if (StartTime != other.StartTime)
                    diffs.Add("Start Time");
                if (EndTime != other.EndTime)
                    diffs.Add("End Time");
            }
            if (Location != other.Location)
                diffs.Add("Location");
            if (IsRepeating != other.IsRepeating)
                diffs.Add("Repeating Event");
            if (IsRepeating && other.IsRepeating)
            {
                if (!RepeatPattern.EquivalentTo(other.RepeatPattern))
                    diffs.Add("Repeat Pattern");
            }
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
    public class RepeatPattern
    {
        private RepeatPattern()
        {
        }
        public static RepeatPattern CreateFromOutlookPattern(RecurrencePattern olPattern)
        {
            RepeatPattern pattern = new RepeatPattern();
            switch (olPattern.RecurrenceType)
            {
                case OlRecurrenceType.olRecursDaily:
                    pattern.IsDaily = true;
                    break;
                case OlRecurrenceType.olRecursMonthly:
                    pattern.IsMonthly = true;
                    pattern.DayOfMonth = olPattern.DayOfMonth;
                    break;
                case OlRecurrenceType.olRecursWeekly:
                    pattern.IsWeekly = true;
                    switch (olPattern.DayOfWeekMask)
                    {
                        case OlDaysOfWeek.olFriday:
                            pattern.DayOfWeek = DayOfWeek.Friday;
                            break;
                        case OlDaysOfWeek.olMonday:
                            pattern.DayOfWeek = DayOfWeek.Monday;
                            break;
                        case OlDaysOfWeek.olSaturday:
                            pattern.DayOfWeek = DayOfWeek.Saturday;
                            break;
                        case OlDaysOfWeek.olSunday:
                            pattern.DayOfWeek = DayOfWeek.Sunday;
                            break;
                        case OlDaysOfWeek.olThursday:
                            pattern.DayOfWeek = DayOfWeek.Thursday;
                            break;
                        case OlDaysOfWeek.olTuesday:
                            pattern.DayOfWeek = DayOfWeek.Tuesday;
                            break;
                        case OlDaysOfWeek.olWednesday:
                            pattern.DayOfWeek = DayOfWeek.Wednesday;
                            break;
                        default:
                            throw new ArgumentException("The outlook pattern cannot be converted into a correct RepeatPattern: occurrs several days a week");
                    }
                    break;
                case OlRecurrenceType.olRecursYearly:
                    pattern.IsYearly = true;
                    pattern.DayOfMonth = olPattern.DayOfMonth;
                    pattern.MonthOfYear = olPattern.MonthOfYear;
                    break;
                default:
                    throw new ArgumentException("Unknown repeat pattern from outlook: " + olPattern.RecurrenceType.ToString());
            }
            pattern.NumRepeats = olPattern.Occurrences;
            pattern.FirstOccurrence = olPattern.PatternStartDate;
            return pattern;
        }

        public static RepeatPattern TryCreateFromOccurrences(IList<DateTime> occurrences)
        {
            RepeatPattern pattern = new RepeatPattern();
            if (occurrences.Count < 2)
                throw new ArgumentException("Cannot calculate repeat pattern unless there's at least two occurrences", "occurrences");
            var local = from occ in occurrences
                        select occ.ToLocalTime();
            List<DateTime> LocalOccurrences = new List<DateTime>(local);
            bool SameMinute = true;
            bool SameHour = true;
            bool SameDayOfMonth = true;
            bool SameDayOfWeek = true;
            bool SameMonth = true;
            // TODO: this interval can vary between occurrances (and still be valid)
            pattern.IntervalTimeLength = LocalOccurrences[1] - LocalOccurrences[0];
            for (int i = 1; i < LocalOccurrences.Count; i++)
            {
                var prev = LocalOccurrences[i - 1];
                var cur = LocalOccurrences[i];
                var diff = cur - prev;
                //if (IntervalTimeLength != diff)
                //  Debugger.Break();
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

            if (!(SameMinute || SameHour || SameDayOfMonth || SameDayOfWeek || SameMonth))
                // No pattern was found
                return null;

            // TODO: Maybe also make sure it's only one of the above
            int intervalDays = (int)Math.Round(pattern.IntervalTimeLength.TotalDays);
            if (intervalDays > 364 && intervalDays < 367)
            {
                // Yearly?
                if (SameMonth && SameDayOfMonth)
                    pattern.IsYearly = true;
                else
                    return null;
            }
            else if (intervalDays < 32 && intervalDays > 27)
            {
                if (SameDayOfMonth)
                    pattern.IsMonthly = true;
                else
                    return null;
            }
            else if (intervalDays == 7)
            {
                if (SameDayOfWeek)
                    pattern.IsWeekly = true;
                else
                    return null;
            }
            else if (intervalDays == 1)
            {
                if (SameHour && SameMinute)
                    pattern.IsDaily = true;
                else
                    return null;
            }
            else
                return null;
            pattern.FirstOccurrence = LocalOccurrences[0];
            pattern.NumRepeats = occurrences.Count;
            pattern.DayOfMonth = pattern.FirstOccurrence.Day;
            pattern.DayOfWeek = pattern.FirstOccurrence.DayOfWeek;
            pattern.MonthOfYear = pattern.FirstOccurrence.Month;            
            return pattern;
        }

        public static RepeatPattern CreateFromOccurrences(IList<DateTime> occurrences)
        {           
            var pattern = TryCreateFromOccurrences(occurrences);
            if (pattern == null)
                throw new ArgumentException("Unable to generate a pattern based on the list of occurrences", "occurrences");
            return pattern;
        }

        
        /// <summary>
        /// Gets the interval between repeats. The unit depends on the IsDaily, IsWeekly etc. properties
        /// </summary>
        public int Interval { get { return 1; } }
        /// <summary>
        /// Gets or sets the length of the interval in actual time.
        /// </summary>
        public TimeSpan IntervalTimeLength { get; set; }
        /// <summary>
        /// Gets or sets the number of occurrences for this repeat pattern
        /// </summary>
        public int NumRepeats { get; set; }
        
        public DayOfWeek DayOfWeek { get; set; }
        public int DayOfMonth { get; set; }
        public int MonthOfYear { get; set; }
        public DateTime FirstOccurrence { get; set; }
        /// <summary>
        /// Gets a value indicating whether this is a yearly event.
        /// </summary>
        public bool IsYearly
        {
            get; set;
        }
        
        /// <summary>
        /// Gets a value indicating whether this is a monthly event.
        /// </summary>
        public bool IsMonthly
        {
            get; set;
        }
        /// <summary>
        /// Gets a value indicating whether this is a weekly event.
        /// </summary>
        public bool IsWeekly
        {
            get; set;
        }

        /// <summary>
        /// Gets a value indicating whether this is a daily event.
        /// </summary>
        public bool IsDaily
        {
            get;
            set;
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current repeat pattern.
        /// </summary>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            if (IsDaily)
                sb.Append("Daily event.");
            if (IsWeekly) 
                sb.AppendFormat("Weekly event, Every {0}.", DayOfWeek);
            if (IsMonthly)
                sb.AppendFormat("Monthly event: On day {0}.", DayOfMonth);
            if (IsYearly)
                sb.AppendFormat("Yearly event: {0:MMMM} {1}.", FirstOccurrence, DayOfMonth);
            if (sb.Length == 0)
                sb.Append("Unknown repeat pattern");
            sb.AppendFormat(" {0} occurrences", NumRepeats);

            return sb.ToString();
        }

        public bool EquivalentTo(RepeatPattern other)
        {
            if (IsYearly != other.IsYearly)
                return false;
            if (IsMonthly != other.IsMonthly)
                return false;
            if (IsWeekly != other.IsWeekly)
                return false;
            if (IsDaily != other.IsDaily)
                return false;
            if ((IsMonthly || IsYearly) && DayOfMonth != other.DayOfMonth)
                return false;
            if (IsWeekly && DayOfWeek != other.DayOfWeek)
                return false;
            if (NumRepeats != other.NumRepeats)
                return false;
            if (Interval != other.Interval)
                return false;
            //if (FirstOccurrence != other.FirstOccurrence)
            //    return false;
            return true;
        }
    }
}
