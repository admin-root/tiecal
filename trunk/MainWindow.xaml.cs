using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace TieCal
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Dependency Properties

        /// <summary>
        /// Gets or sets a value indicating whether this instance is busy working with calendar synchronization. This is a dependency property.
        /// </summary>
        [Description("Gets or sets a value indicating whether this instance is busy working with calendar synchronization.")]
        public bool IsSynchronizing
        {
            get { return (bool)GetValue(IsSynchronizingProperty); }
            set { SetValue(IsSynchronizingProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsWorking.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsSynchronizingProperty =
            DependencyProperty.Register("IsSynchronizing", typeof(bool), typeof(MainWindow), new UIPropertyMetadata(false, new PropertyChangedCallback(IsSynchronizingProperty_Changed)));

        /// <summary>
        /// Gets or sets a value indicating whether to run in simulation mode (no changes written to any calendar). This is a dependency property.
        /// </summary>
        [Description("Gets or sets a value indicating whether to run in simulation mode (no changes written to any calendar).")]
        public bool DryRun
        {
            get { return (bool)GetValue(DryRunProperty); }
            set { SetValue(DryRunProperty, value); }
        }

        // Using a DependencyProperty as the backing store for DryRun.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DryRunProperty =
            DependencyProperty.Register("DryRun", typeof(bool), typeof(MainWindow), new UIPropertyMetadata(false));


        #endregion
        #region Routed Events
        public static RoutedEvent SynchronizationStartedEvent = EventManager.RegisterRoutedEvent("SynchronizationStarted", RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MainWindow));
        public static RoutedEvent SynchronizationEndedEvent = EventManager.RegisterRoutedEvent("SynchronizationEnded", RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MainWindow));
        /// <summary>
        /// Occurs when synchronization is started.
        /// </summary>
        [Description("Occurs when synchronization is started.")]
        public event RoutedEventHandler SynchronizationStarted
        {
            add { AddHandler(SynchronizationStartedEvent, value); }
            remove { RemoveHandler(SynchronizationStartedEvent, value); }
        }
        /// <summary>
        /// Occurs when synchronization has ended.
        /// </summary>
        [Description("Occurs when synchronization has ended.")]
        public event RoutedEventHandler SynchronizationEnded
        {
            add { AddHandler(SynchronizationEndedEvent, value); }
            remove { RemoveHandler(SynchronizationEndedEvent, value); }
        }
        #endregion

        public static void IsSynchronizingProperty_Changed(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            MainWindow syncWindow = (MainWindow)sender;

            if ((bool)e.NewValue)
                syncWindow.RaiseEvent(new RoutedEventArgs(SynchronizationStartedEvent));
            else
                syncWindow.RaiseEvent(new RoutedEventArgs(SynchronizationEndedEvent));
        }

        private NotesReader _notesReader;
        private OutlookManager _outlookManager;
        private ProgramSettings settings;
        
        public MainWindow()
        {
            InitializeComponent();
            settings = ProgramSettings.LoadSettings();
            txtNotesPassword.Password = settings.NotesPassword;            
            _notesReader = new NotesReader();
            _outlookManager = new OutlookManager();
            _notesReader.FetchCalendarWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(notesworker_RunWorkerCompleted);
            _notesReader.FetchCalendarWorker.ProgressChanged += new ProgressChangedEventHandler(notesworker_ProgressChanged);
            _outlookManager.FetchCalendarWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(outlookworker_RunWorkerCompleted);
            _outlookManager.FetchCalendarWorker.ProgressChanged += new ProgressChangedEventHandler(outlookworker_ProgressChanged);

            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (settings.NotesPassword != null && settings.NotesPassword.Length > 0)
                RefreshNotesDatabases();
            else
                expSettings.IsExpanded = true;
            DryRun = settings.DryRun;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            settings.DryRun = DryRun;
            settings.Save();
            base.OnClosing(e);
        }

        private void BeginFetchCalendarEntries()
        {
            IsSynchronizing = true;
            txtStatusMessage.Text = "Reading calendars";
            _notesReader.BeginFetchCalendarEntries();
            _outlookManager.BeginFetchCalendarEntries();
        }

        void outlookworker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbarOutlook.Value = e.ProgressPercentage;
        }

        void notesworker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbarNotes.Value = e.ProgressPercentage;
        }
        private bool fetchFailed = false;
        private int doneCount;

        void outlookworker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show("Outlook calendar fetch failed: " + e.Error.Message);
                fetchFailed = true;
            }
            if (e.Cancelled)
                fetchFailed = true;
            doneCount++;
            if (doneCount == 2)
                MergeCalendarEntries();
        }

        void notesworker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show("Notes calendar fetch failed: " + e.Error.Message);
                fetchFailed = true;
            }
            if (e.Cancelled)
                fetchFailed = true;
            doneCount++;
            if (doneCount == 2)
                MergeCalendarEntries();
        }
        
        private void MergeCalendarEntries()
        {
            try
            {
                // TODO: This should probably move to MergeWindow or a separate class
                txtStatusMessage.Text = "Processing Calendar Entries...";
                if (fetchFailed)
                    return;
                var mapping = new EntryIDMapping();
                try
                {
                    mapping.Load();
                    mapping.Cleanup(_notesReader.CalendarEntries, _outlookManager.CalendarEntries);
                }
                catch (System.IO.FileNotFoundException) { }
                var lowerLimit = DateTime.Now - TimeSpan.FromDays(30);
                var upperLimit = DateTime.Now + TimeSpan.FromDays(30);
                var entriesToMerge = from calEntry in _notesReader.CalendarEntries
                                     where calEntry.IsRepeating == false &&
                                     calEntry.StartTime > lowerLimit && calEntry.EndTime < upperLimit
                                     select calEntry;
                
                foreach (var calEntry in entriesToMerge)
                    calEntry.OutlookID = mapping.GetOutlookID(calEntry.NotesID);                
                foreach (var calEntry in _outlookManager.CalendarEntries)
                    calEntry.NotesID = mapping.GetNotesID(calEntry.OutlookID);

                var newEntries = from notesEntry in entriesToMerge
                                 where !(from outlookEntry in _outlookManager.CalendarEntries select outlookEntry.NotesID).Contains(notesEntry.NotesID)
                                 select notesEntry;
                var changedEntries = from notesEntry in entriesToMerge
                                     join outlookEntry in _outlookManager.CalendarEntries on notesEntry.OutlookID equals outlookEntry.OutlookID
                                     where notesEntry.OutlookID == outlookEntry.OutlookID && notesEntry.DiffersFrom(outlookEntry)
                                     select new { Entry = notesEntry, Differences = notesEntry.GetDifferences(outlookEntry) };

                var oldEntries = from outlookEntry in _outlookManager.CalendarEntries
                                 where !(from notesEntry in entriesToMerge select notesEntry.OutlookID).Contains(outlookEntry.OutlookID)
                                 select outlookEntry;
                MergeWindow mergeWin = new MergeWindow(newEntries, changedEntries, oldEntries);
                bool doMerge = (mergeWin.ShowDialog() == true);
                if (doMerge && !DryRun)
                {
                    _outlookManager.RemoveCalendarEntries(oldEntries);
                    List<CalendarEntry> changedEntries2 = new List<CalendarEntry>();
                    foreach (var chg in changedEntries)
                        changedEntries2.Add(chg.Entry);
                    _outlookManager.MergeCalendarEntries(changedEntries2);
                    _outlookManager.AddCalendarEntries(newEntries);
                    mapping.AddRange(entriesToMerge);
                    mapping.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to merge: " + ex.Message);
            }
            finally
            {
                txtStatusMessage.Text = "All Done";
                IsSynchronizing = false;
            }
        }
        private void RefreshNotesDatabases()
        {
            _notesReader.Password = settings.NotesPassword;
            cmbNotesDB.ItemsSource = _notesReader.GetAvailableDatabases();
            if (settings.NotesDatabase != null)
                cmbNotesDB.SelectedItem = settings.NotesDatabase;
        }
        private void btnSync_Click(object sender, RoutedEventArgs e)
        {
            doneCount = 0;
            fetchFailed = false;
            _notesReader.Password = settings.NotesPassword;
            _notesReader.DatabaseFile = settings.NotesDatabase;
            BeginFetchCalendarEntries();
        }

        private void btnCancelSync_Click(object sender, RoutedEventArgs e)
        {
            if (_notesReader.FetchCalendarWorker.IsBusy)
                _notesReader.FetchCalendarWorker.CancelAsync();
            if (_outlookManager.FetchCalendarWorker.IsBusy)
                _outlookManager.FetchCalendarWorker.CancelAsync();
        }

        private void txtNotesPassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
            settings.NotesPassword = txtNotesPassword.Password;
        }

        private void btnRefreshNotesDB_Click(object sender, RoutedEventArgs e)
        {
            RefreshNotesDatabases();
        }

        private void cmbNotesDB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            settings.NotesDatabase = (string) cmbNotesDB.SelectedItem;
        }
    }

    /// <summary>
    /// Negates a boolean value
    /// </summary>
    public class BoolToOppositeBoolConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        #endregion
    }
}
