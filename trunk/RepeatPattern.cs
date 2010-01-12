using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace TieCal
{
    /// <summary>
    /// Analyzer for a list of <see cref="DateTime"/> objects which tries to find a pattern that can be applied when creating/updating outlook entries
    /// </summary>
    public class RepeatPattern
    {
        private RepeatPattern()
        {
            Interval = 1;
        }
        public static RepeatPattern CreateFromOutlookPattern(RecurrencePattern olPattern)
        {
            RepeatPattern pattern = new RepeatPattern();
            switch (olPattern.RecurrenceType)
            {
                case OlRecurrenceType.olRecursDaily:
                    pattern.IsDaily = true;
                    pattern.Interval = olPattern.Interval;
                    break;
                case OlRecurrenceType.olRecursMonthly:
                    pattern.IsMonthly = true;
                    pattern.DayOfMonth = olPattern.DayOfMonth;
                    pattern.Interval = olPattern.Interval;
                    break;
                case OlRecurrenceType.olRecursWeekly:
                    pattern.IsWeekly = true;
                    pattern.Interval = olPattern.Interval;
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
                            // It's multiple times per week
                            pattern.IsWeekly = false;
                            pattern.IsWeeklyMultipleDays = true;
                            pattern.DayOfWeekMask = (DaysOfWeek)olPattern.DayOfWeekMask;
                            break;
                        //throw new ArgumentException("The outlook pattern cannot be converted into a correct RepeatPattern: occurrs several days a week");
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

        private static List<DateTime> GetLocalTimes(IEnumerable<DateTime> occurrences)
        {
            var local = from occ in occurrences
                        select occ.ToLocalTime();
            return new List<DateTime>(local);
        }

        public static RepeatPattern TryCreateFromOccurrences(IList<DateTime> occurrences)
        {
            RepeatPattern pattern = new RepeatPattern();
            if (occurrences.Count < 2)
                throw new ArgumentException("Cannot calculate repeat pattern unless there's at least two occurrences", "occurrences");
            List<DateTime> LocalOccurrences = GetLocalTimes(occurrences);
            bool SameMinute = true;
            bool SameHour = true;
            bool SameDayOfMonth = true;
            bool SameDayOfWeek = true;
            bool SameMonth = true;
            // TODO: this interval can vary between occurrances (and still be valid)
            pattern.IntervalTimeLength = LocalOccurrences[1] - LocalOccurrences[0];
            var intervalDayRange = new IntervalRange(7);
            var intervalMonthRange = new IntervalRange(12);
            for (int i = 1; i < LocalOccurrences.Count; i++)
            {
                var prev = LocalOccurrences[i - 1];
                var cur = LocalOccurrences[i];
                int monthDiff = (cur.Year - prev.Year) * 12 + (cur.Month - prev.Month);
                var diff = cur - prev;
                intervalDayRange.Add((int)Math.Round(diff.TotalDays));
                intervalMonthRange.Add(monthDiff);
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
            // Fun part.. try to make sense out of all this. Would be nicer if Notes just told us the repeat pattern, but hey.. it's a piece of crap so what can you expect?
            // TODO: Re-factor this logic a bit.. way too many exit points
            if (intervalDayRange.ItemsAreIdentical)
            {
                // We have the same amount of days between each occurrence. This is either a daily or weekly event with one occurrence per day/week
                int days = intervalDayRange.First();
                if (days < 7)
                {
                    if (SameHour && SameMinute)
                    {
                        pattern.IsDaily = true;
                        pattern.Interval = days;
                    }
                    else
                        return null;
                }
                else
                {
                    // Weekly or Bi-, Tri-, ..., -weekly
                    // Actually, this is just a special case of several-days-a-week pattern below
                    int weeksBetween = days / 7;
                    if (SameDayOfWeek && weeksBetween < 10)
                    {
                        pattern.IsWeekly = true;
                        pattern.Interval = weeksBetween;
                    }
                    else
                        return null;
                }
            }
            else if (intervalMonthRange.ItemsAreIdentical)
            {
                // Same amount of months between each occurrence. This is monthly or yearly
                int months = intervalMonthRange.First();
                if (months == 12)
                {
                    // Yearly?
                    if (SameMonth && SameDayOfMonth)
                    {
                        pattern.IsYearly = true;
                    }
                    else
                        return null;
                }
                else
                {
                    // Monthly?
                    if (SameDayOfMonth && months < 12)
                    {
                        pattern.IsMonthly = true;
                        pattern.Interval = months;
                    }
                    else
                        return null;
                }
            }
            else if (intervalDayRange.HasRepeatingCycle)
            {
                // Weird repeating pattern, e.g. several times a week
                var intervalCycle = intervalDayRange.GetRepeatingCycle();
                var cycleLength = LocalOccurrences[intervalCycle.Count] - LocalOccurrences[0];
                if ((int)cycleLength.TotalDays % 7 == 0)
                {
                    // It's weekly (or bi-weekly etc.), but several days a week
                    int weeksBetween = (int)cycleLength.TotalDays / 7;
                    if (SameHour && SameMinute && weeksBetween < 10)
                    {
                        var dayMask = GetDaysOfWeek(LocalOccurrences);
                        if (dayMask == DaysOfWeek.None)
                            return null;
                        pattern.IsWeeklyMultipleDays = true;
                        pattern.DayOfWeekMask = dayMask;
                        pattern.Interval = weeksBetween;
                    }
                    else
                        return null;
                }
                else
                    return null;
            }
            else
                return null;
            pattern.FirstOccurrence = LocalOccurrences[0];
            pattern.LastOccurrence = LocalOccurrences[LocalOccurrences.Count - 1];
            pattern.NumRepeats = occurrences.Count;
            pattern.DayOfMonth = pattern.FirstOccurrence.Day;
            pattern.DayOfWeek = pattern.FirstOccurrence.DayOfWeek;
            pattern.MonthOfYear = pattern.FirstOccurrence.Month;
            return pattern;
        }

        /// <summary>
        /// Checks to see which days of the week the repeating event occurrs on. If an answer can't be found, DaysOfWeek.None is returned
        /// </summary>
        private static DaysOfWeek GetDaysOfWeek(IList<DateTime> occurrences)
        {
            DaysOfWeek days = DaysOfWeek.None;
            var dayList = new List<DayOfWeek>();
            bool valid = false;
            foreach (var occurrence in occurrences)
            {
                if (dayList.Count > 0 && dayList[0] == occurrence.DayOfWeek)
                {
                    // The list contains all days of a single week now, proceed to verify
                    valid = true;
                    break;
                }
                dayList.Add(occurrence.DayOfWeek);
            }
            if (!valid)
                return DaysOfWeek.None;
            for (int i = 0; i < occurrences.Count; i++)
            {
                if (occurrences[i].DayOfWeek != dayList[i % dayList.Count])
                    return DaysOfWeek.None;
                switch (occurrences[i].DayOfWeek)
                {
                    case DayOfWeek.Friday:
                        days |= DaysOfWeek.Friday;
                        break;
                    case DayOfWeek.Monday:
                        days |= DaysOfWeek.Monday;
                        break;
                    case DayOfWeek.Saturday:
                        days |= DaysOfWeek.Saturday;
                        break;
                    case DayOfWeek.Sunday:
                        days |= DaysOfWeek.Sunday;
                        break;
                    case DayOfWeek.Thursday:
                        days |= DaysOfWeek.Thursday;
                        break;
                    case DayOfWeek.Tuesday:
                        days |= DaysOfWeek.Tuesday;
                        break;
                    case DayOfWeek.Wednesday:
                        days |= DaysOfWeek.Wednesday;
                        break;
                    default:
                        break;
                }
            }
            return days;
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
        public int Interval { get; set; }
        /// <summary>
        /// Gets or sets the length of the interval in actual time.
        /// </summary>
        public TimeSpan IntervalTimeLength { get; set; }
        /// <summary>
        /// Gets or sets the number of occurrences for this repeat pattern
        /// </summary>
        public int NumRepeats { get; set; }

        public DayOfWeek DayOfWeek { get; set; }
        public DaysOfWeek DayOfWeekMask { get; set; }
        public int DayOfMonth { get; set; }
        public int MonthOfYear { get; set; }
        public DateTime FirstOccurrence { get; set; }
        public DateTime LastOccurrence { get; set; }
        /// <summary>
        /// Gets a value indicating whether this is a yearly event.
        /// </summary>
        public bool IsYearly
        {
            get;
            set;
        }

        /// <summary>
        /// Gets a value indicating whether this is a monthly event. This means it occurrs once per month on a specific day (1-31)
        /// </summary>
        public bool IsMonthly
        {
            get;
            set;
        }
        /// <summary>
        /// Gets a value indicating whether this is a weekly event. This means it occurrs once per week on a specific weekday (monday - sunday)
        /// </summary>
        public bool IsWeekly
        {
            get;
            set;
        }

        /// <summary>
        /// Gets a value indicating whether this is a weekly event with multiple days. This means that it occurrs several times per week but on the same days each week (monday - sunday)
        /// </summary>
        public bool IsWeeklyMultipleDays
        {
            get;
            set;
        }

        /// <summary>
        /// Gets a value indicating whether this is a daily event. This means it occurrs every single day at the same time
        /// </summary>
        public bool IsDaily
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the ordinal (1st, 2nd, 3rd, 4th...) string representation of the specified number
        /// </summary>
        /// <param name="number">The number to represent ordinally</param>
        private string GetOrdinalString(int number)
        {
            switch (Interval)
            {
                case 1:
                    return "1st";
                case 2:
                    return "2nd";
                case 3:
                    return "3rd";
                default:
                    return String.Format("{0}th", number);
            }
        }
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current repeat pattern.
        /// </summary>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder("Every ");
            if (Interval > 1)
            {
                sb.Append(GetOrdinalString(Interval));
                sb.Append(' ');
            }
            if (IsDaily)
                sb.Append("day");
            if (IsWeekly)
                sb.AppendFormat("week on {0}.", DayOfWeek);
            if (IsWeeklyMultipleDays)
                sb.AppendFormat("week on {0}.", DayOfWeekMask);
            if (IsMonthly)
                sb.AppendFormat("month on the {0}.", GetOrdinalString(DayOfMonth));
            if (IsYearly)
                sb.AppendFormat("year, on {0:MMMM} {1}.", FirstOccurrence, GetOrdinalString(DayOfMonth));
            if (sb.Length == 0)
                sb.Append("Unknown repeat pattern");
            if (NumRepeats != 1)
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
    /// <summary>
    /// Provides an enum representing the different days of the week. The values can be bitwise or'ed together and maps directly to 
    /// Outlooks olDaysOfWeek enumeration
    /// </summary>
    [Flags]
    public enum DaysOfWeek
    {
        None = 0,
        Sunday = 1,
        Monday = 2,
        Tuesday = 4,
        Wednesday = 8,
        Thursday = 16,
        Friday = 32,
        Saturday = 64,
    }

    internal class IntervalRange
    {
        private List<int> items = new List<int>();
        private int _repeatSumLimit = Int32.MaxValue;
        public IntervalRange(int repeatSumLimit)
        {
            _repeatSumLimit = repeatSumLimit;
        }

        public void Add(int range)
        {
            items.Add(range);
        }

        public bool AddUnique(int range)
        {
            // Check if it exists already
            foreach (int i in items)
                if (i == range)
                    return false;
            items.Add(range);
            return true;
        }

        private List<int> GetRepeatingCycleInternal()
        {
            if (items.Count < 2)
                return null;
            if (GetBiggestDiff() == 0)
                // All values are the same
                return null;
            // Brute force naive way of finding cycles, would be neat to use Floyd's cycle-finding algorithm instead ( http://en.wikipedia.org/wiki/Cycle_detection#Tortoise_and_hare )
            // but that doesn't seem to work here (since it will find the shortest cycle, we want the longest)
            var cycle = new List<int>();
            bool found = false;
            for (int i = 0;
                i < items.Count - 1;
                i++)
            {
                var cur = items[i];
                var next = items[i + 1];
                if (cycle.Count > 1 && cycle[0] == cur)
                {
                    // We've found the start of our proposed cycle now (but we're not sure it's the end yet)
                    if (cycle[1] == next && cycle.Sum() >= _repeatSumLimit)
                    {
                        found = true;
                        break;
                    }
                }
                cycle.Add(cur);
            }
            if (!found ||                         // No pattern found
                cycle.Count == items.Count     // All items added to pattern
                )   // Pattern doesn't fully repeat throughout the interval range
                return null;
            for (int i = 0; i < items.Count; i++)
            {
                if (cycle[i % cycle.Count] != items[i])
                    return null;
            }
            return cycle;
        }

        public IList<int> GetRepeatingCycle()
        {
            var pattern = GetRepeatingCycleInternal();
            if (pattern == null)
                throw new InvalidOperationException("No repeating pattern could be found for the intervals");

            return pattern;
        }

        /// <summary>
        /// Gets a value indicating whether there is a repeating pattern in the intervals, for instance {1,2,3, 1,2,3, 1,2,3} or {1,5,2, 1,5,2, ...}.
        /// </summary>
        /// <remarks>
        /// Special Cases:
        /// * If all intervals are identical, this property is true.
        /// * If there are zero (0) or one (1) item in the list, then this property is false.
        /// </remarks>
        /// <value>
        /// 	<c>true</c> if the intervals has a repeating pattern; otherwise, <c>false</c>.
        /// </value>
        public bool HasRepeatingCycle
        {
            get
            {
                var pattern = GetRepeatingCycleInternal();

                if (pattern != null)
                    return true;

                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether all intervals in the list are identical.
        /// </summary>
        public bool ItemsAreIdentical
        {
            get
            {
                if (items.Count == 0)
                    return false;
                if (items.Count == 1)
                    return true;
                for (int i = 1; i < items.Count; i++)
                {
                    if (items[i] != items[i - 1])
                        return false;
                }
                return true;
            }
        }

        public int GetMax()
        {
            if (items.Count == 0)
                throw new InvalidOperationException("No items in list");
            var max = Int32.MinValue;
            foreach (int i in items)
                if (i > max)
                    max = i;
            return max;
        }

        public int GetMin()
        {
            if (items.Count == 0)
                throw new InvalidOperationException("No items in list");
            var min = Int32.MaxValue;
            foreach (int i in items)
                if (i < min)
                    min = i;
            return min;
        }

        public int GetAverage()
        {
            if (items.Count == 0)
                throw new InvalidOperationException("No items in list");
            var avg = 0;
            foreach (int i in items)
                avg += i;
            return avg / items.Count;
        }

        public int GetBiggestDiff()
        {
            return GetMax() - GetMin();
        }

        public int GetBiggestDiffFromAverage()
        {
            var avg = GetAverage();
            var upDiff = GetMax() - avg;
            var downDiff = avg - GetMin();

            return Math.Max(upDiff, downDiff);
        }

        public double GetBiggestDiffPercentage()
        {
            return (double)GetBiggestDiffFromAverage() / (double)GetAverage();
        }

        /// <summary>
        /// Gets a value indicating whether there are a difference between the items in the dayrange (true means all items are equal, false means there is at least one that is different from the others).
        /// </summary>
        public bool HasDiff
        {
            get
            {
                return items.Count < 2 || GetBiggestDiff() == 0;
            }
        }

        public int First()
        {
            return items[0];
        }
    }
}
