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

        int NumberOfSkippedEntries { get; }
        /// <summary>
        /// Gets the background worker used to fetch calendar entries in the background.
        /// </summary>
        BackgroundWorker FetchCalendarWorker { get; }
    }
}
