using System;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Navigation;

namespace TieCal
{
	public partial class SelectNotesDbDialog
	{
		public SelectNotesDbDialog()
		{
			this.InitializeComponent();
		}

        public SelectNotesDbDialog(NotesReader notesReader) : this()
        {
            RefreshNotesDatabases(notesReader);
        }

        private void RefreshNotesDatabases(NotesReader notesReader)
        {
            cmbNotesDB.Items.Clear();
            cmbNotesDB.ItemsSource = notesReader.GetAvailableDatabases();
            if (ProgramSettings.Instance.NotesDatabase != null)
                cmbNotesDB.SelectedItem = ProgramSettings.Instance.NotesDatabase;
            else
            {
                // Make a default selection. The one with the calendar is most often the one named: mail\<username>.nsf
                foreach (var item in cmbNotesDB.Items)
                {
                    if (item.ToString().StartsWith(@"mail\") && item.ToString().EndsWith(".nsf"))
                    {
                        cmbNotesDB.SelectedItem = item;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the notes database that the user selected.
        /// </summary>
        public string NotesDatabase
        {
            get { return (string)cmbNotesDB.SelectedItem; }
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            //Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            //Close();
        }
	}
}