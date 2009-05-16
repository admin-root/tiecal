using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace TieCal
{
    public interface ICalendarReader
    {
        /// <summary>
        /// Gets or sets the calendar entries. 
        /// </summary>
        /// <value>The calendar entries.</value>
        IEnumerable<CalendarEntry> CalendarEntries { get; }
        /// <summary>
        /// Gets the version of the calendar application that this Calendar Reader operates against.
        /// </summary>
        /// <value>The version reported from the calendar application.</value>
        string CalendarAppVersion { get; }
        /// <summary>
        /// Gets the number of calendar entries that was skipped while reading the calendar.
        /// </summary>
        /// <value>The number of skipped entries.</value>
        int NumberOfSkippedEntries { get; }
        /// <summary>
        /// Gets the background worker used to fetch calendar entries in the background.
        /// </summary>
        BackgroundWorker FetchCalendarWorker { get; }
    }
}
