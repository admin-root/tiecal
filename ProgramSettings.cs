using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Security;

namespace TieCal
{
    /// <summary>
    /// Holds settings for TieCal
    /// </summary>
    [Serializable]
    public class ProgramSettings
    {
        public ProgramSettings() 
        {
            DryRun = true;
        }
        /// <summary>
        /// Gets the filename where settings are saved.
        /// </summary>
        private static string SaveFilename 
        {
            get
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TieCal");
                Directory.CreateDirectory(folder);
                return Path.Combine(folder, "ProgramSettings.txt");
            }
        }

        public static ProgramSettings LoadSettings()
        {
            try
            {
                using (TextReader writer = new StreamReader(SaveFilename))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(ProgramSettings));
                    return (ProgramSettings)serializer.Deserialize(writer);
                }
            }
            catch (FileNotFoundException)
            {
                return new ProgramSettings();
            }
        }

        public void Save()
        {
            using (TextWriter writer = new StreamWriter(SaveFilename))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ProgramSettings));
                serializer.Serialize(writer, this);
            }
        }
        public string NotesDatabase { get; set; }
        public string NotesPassword { get; set; }
        public bool DryRun { get; set; }
    }
}
