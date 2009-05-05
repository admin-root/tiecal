using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace TieCal
{
    /// <summary>
    /// Contains mapping between Notes ID and Outlook ID
    /// </summary>
    public class EntryIDMapping
    {
        private Dictionary<string, string> notesToOutlook = new Dictionary<string,string>();
        private Dictionary<string, string> outlookToNotes = new Dictionary<string,string>();        
        private string Filename { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="EntryIDMapping"/> class.
        /// </summary>
        public EntryIDMapping()
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TieCal");
            Directory.CreateDirectory(folder);
            Filename = Path.Combine(folder, "IDMapping.txt");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EntryIDMapping"/> class.
        /// </summary>
        /// <param name="entries">The entries from which to add initial ID mapping information.</param>
        public EntryIDMapping(IEnumerable<CalendarEntry> entries)
            : this()
        {
            foreach (var entry in entries)
            {
                if (entry.NotesID != null && entry.OutlookID != null)
                    AddPair(entry.NotesID, entry.OutlookID);
            }
        }

        /// <summary>
        /// Adds the ID pair to the store.
        /// </summary>
        /// <param name="notesID">The notes ID.</param>
        /// <param name="outlookID">The outlook ID.</param>
        public void AddPair(string notesID, string outlookID)
        {
            if (notesToOutlook.ContainsKey(notesID))
            {
                outlookToNotes[outlookID] = notesID;
                notesToOutlook[notesID] = outlookID;
            }
            else
            {
                outlookToNotes.Add(outlookID, notesID);
                notesToOutlook.Add(notesID, outlookID);
            }
        }

        /// <summary>
        /// Removes the ID pair from memory.
        /// </summary>
        /// <param name="notesID">The notes ID.</param>
        /// <param name="outlookID">The outlook ID.</param>
        private void RemovePair(string notesID, string outlookID)
        {
            if (outlookToNotes.ContainsKey(outlookID))
                outlookToNotes.Remove(outlookID);
            if (notesToOutlook.ContainsKey(notesID))
                notesToOutlook.Remove(notesID);
        }

        private void RemoveNotesID(string notesID)
        {
            if (notesToOutlook.ContainsKey(notesID))
                notesToOutlook.Remove(notesID);
            var qry = from pair in outlookToNotes
                      where pair.Value == notesID
                      select pair.Key;
            if (qry.Count() > 0)
                outlookToNotes.Remove(qry.First());
        }

        private void RemoveOutlookID(string outlookID)
        {
            if (outlookToNotes.ContainsKey(outlookID))
                outlookToNotes.Remove(outlookID);

            var qry = from pair in notesToOutlook
                      where pair.Value == outlookID
                      select pair.Key;
            if (qry.Count() > 0)
                outlookToNotes.Remove(qry.First());
        }

        /// <summary>
        /// Clears all mappings.
        /// </summary>
        public void Clear()
        {
            notesToOutlook.Clear();
            outlookToNotes.Clear();
        }

        /// <summary>
        /// Saves the mapping to disk.
        /// </summary>
        public void Save()
        {
            using (TextWriter writer = new StreamWriter(Filename))
            {
                writer.WriteLine("# NotesID:OutlookID");
                foreach (var entry in notesToOutlook)
                    writer.WriteLine("{0}:{1}",entry.Key, entry.Value);
            }
        }

        /// <summary>
        /// Loads previously saved mapping from disk.
        /// </summary>
        public void Load()
        {
            using (TextReader reader = new StreamReader(Filename))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith("#"))
                        continue;
                    string[] pieces = line.Split(':');
                    AddPair(pieces[0], pieces[1]);
                }
            }
        }

        /// <summary>
        /// Gets the notes ID for the corresponding outlook ID, or null if it doesn't exist.
        /// </summary>
        /// <param name="outlookID">The outlook ID.</param>
        /// <returns></returns>
        public string GetNotesID(string outlookID)
        {
            if (outlookToNotes.ContainsKey(outlookID))
                return outlookToNotes[outlookID];
            return null;
        }

        /// <summary>
        /// Gets the outlook ID for the corresponding notes ID, or null if it doesn't exist.
        /// </summary>
        /// <param name="notesID">The notes ID.</param>
        /// <returns></returns>
        public string GetOutlookID(string notesID)
        {
            if (notesToOutlook.ContainsKey(notesID))
                return notesToOutlook[notesID];
            return null;
        }

        public void AddRange(IEnumerable<CalendarEntry> calendarEntries)
        {
            foreach (var entry in calendarEntries)
            {
                AddPair(entry.NotesID, entry.OutlookID);
            }
        }

        public IEnumerable<string> KnownNotesIDs
        {
            get { return notesToOutlook.Keys; }
        }

        public IEnumerable<string> KnownOutlookIDs
        {
            get { return notesToOutlook.Values; }
        }

        /// <summary>
        /// Removes all mapping information that is no longer available in the two specified collection of entries
        /// </summary>
        /// <param name="notesEntries">The notes entries.</param>
        /// <param name="outlookEntries">The outlook entries.</param>
        public void Cleanup(IEnumerable<CalendarEntry> notesEntries, IEnumerable<CalendarEntry> outlookEntries)
        {
            if (notesEntries != null && notesEntries.Count() > 0)
            {
                var notesIDsToRemove = from id in KnownNotesIDs
                                       where !(from entry in notesEntries select entry.NotesID).Contains(id)
                                       select id;
                List<string> notesIds = new List<string>(notesIDsToRemove.AsEnumerable());
                foreach (string id in notesIds)
                    RemoveNotesID(id);
            }
            if (outlookEntries != null && outlookEntries.Count() > 0)
            {
                var outlookIDsToRemove = from id in KnownOutlookIDs
                                         where !(from entry in outlookEntries select entry.OutlookID).Contains(id)
                                         select id;
                List<string> outlookIds = new List<string>(outlookIDsToRemove.AsEnumerable());
                foreach (string id in outlookIds)
                    RemoveOutlookID(id);
            }
        }
    }
}
