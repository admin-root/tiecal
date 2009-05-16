using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
namespace TieCal
{
    class CalendarMerger
    {
        EntryIDMapping mapping = new EntryIDMapping();
        public CalendarMerger()
        {
            try
            {
                mapping.Load();
            }
            catch (System.IO.FileNotFoundException) { }
            Worker = new BackgroundWorker();
            Worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            Worker.WorkerReportsProgress = true;
            Worker.WorkerSupportsCancellation = true;
        }

        void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Worker.ReportProgress(0);
            mapping.Cleanup(NotesEntries, OutlookEntries);
            if (Worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
            Worker.ReportProgress(10);
            var lowerLimit = DateTime.Now - TimeSpan.FromDays(30);
            var upperLimit = DateTime.Now + TimeSpan.FromDays(30);
            var entriesToMerge = from calEntry in NotesEntries
                                 where calEntry.IsRepeating == false && calEntry.OccursInInterval(lowerLimit, upperLimit)
                                 select calEntry;
            if (Worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
            Worker.ReportProgress(20);
            foreach (var calEntry in entriesToMerge)
                calEntry.OutlookID = mapping.GetOutlookID(calEntry.NotesID);
            foreach (var calEntry in OutlookEntries)
                calEntry.NotesID = mapping.GetNotesID(calEntry.OutlookID);
            if (Worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
            Worker.ReportProgress(25);
            
            var newEntries = from notesEntry in entriesToMerge
                             where !(from outlookEntry in OutlookEntries select outlookEntry.NotesID).Contains(notesEntry.NotesID)
                             select new ModifiedEntry(notesEntry, Modification.New);
            if (Worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
            Worker.ReportProgress(40);

            var changedEntries = from notesEntry in entriesToMerge
                                 join outlookEntry in OutlookEntries on notesEntry.OutlookID equals outlookEntry.OutlookID
                                 where notesEntry.OutlookID == outlookEntry.OutlookID && notesEntry.DiffersFrom(outlookEntry)
                                 select new ModifiedEntry(notesEntry, Modification.Modified, notesEntry.GetDifferences(outlookEntry));
            if (Worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
            Worker.ReportProgress(60);

            var oldEntries = from outlookEntry in OutlookEntries
                             where !(from notesEntry in entriesToMerge select notesEntry.OutlookID).Contains(outlookEntry.OutlookID)
                             select new ModifiedEntry(outlookEntry, Modification.Removed);
            if (Worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
            Worker.ReportProgress(80);

            ModifiedEntries = new List<ModifiedEntry>(newEntries.Count() + changedEntries.Count() + oldEntries.Count());
            ModifiedEntries.AddRange(newEntries);
            ModifiedEntries.AddRange(changedEntries);
            ModifiedEntries.AddRange(oldEntries);
            e.Result = ModifiedEntries;
            Worker.ReportProgress(100);
        }

        public IEnumerable<CalendarEntry> NotesEntries { get; set; }
        public IEnumerable<CalendarEntry> OutlookEntries { get; set; }
        public List<ModifiedEntry> ModifiedEntries { get; set; }
        public BackgroundWorker Worker { get; set; }

        /// <summary>
        /// Saves the notes/outlook ID mappings. Call this method after entries have been saved to outlook because that
        /// is when they get an ID assigned.
        /// </summary>
        public void SaveMappings()
        {
            foreach (var entry in ModifiedEntries)
                if (entry.ApplyModification && entry.Entry.OutlookID != null && entry.Entry.NotesID != null)
                    mapping.AddPair(entry.Entry.NotesID, entry.Entry.OutlookID);
            mapping.Save();
        }
    }
}
