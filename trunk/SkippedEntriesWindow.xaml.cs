using System;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Collections.Generic;
using System.Text;

namespace TieCal
{
	public partial class SkippedEntriesWindow
	{
		public SkippedEntriesWindow()
		{
			this.InitializeComponent();
			
			// Insert code required on object creation below this point.
		}
        public void Show(IEnumerable<SkippedEntry> skippedEntries)
        {
            lstEntries.ItemsSource = skippedEntries;
            this.Show();
        }
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnClipboard_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            foreach (SkippedEntry entry in lstEntries.ItemsSource)
            {
                sb.AppendFormat("{0};{1};{2};{3};{4}", entry.Reason, entry.CalendarEntry.Subject, entry.CalendarEntry.StartTimeLocal, entry.CalendarEntry.EndTimeLocal, entry.CalendarEntry.Location);
                sb.AppendLine();
            }
            Clipboard.SetText(sb.ToString());
            MessageBox.Show("List of entries successfully copied to clipboard");
        }
	}
}