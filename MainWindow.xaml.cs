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

        /// <summary>
        /// Gets or sets a value indicating whether this instance is ready to synchronize the calendars.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is ready to synchronize; otherwise, <c>false</c>.
        /// </value>
        public bool IsReadyToSynchronize
        {
            get { return (bool)GetValue(IsReadyToSynchronizeProperty); }
            set { SetValue(IsReadyToSynchronizeProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsReadyToSynchronize.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsReadyToSynchronizeProperty =
            DependencyProperty.Register("IsReadyToSynchronize", typeof(bool), typeof(MainWindow), new UIPropertyMetadata(false, new PropertyChangedCallback(IsReadyToSynchronizeProperty_Changed)));

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
        private static void IsReadyToSynchronizeProperty_Changed(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            MainWindow syncWindow = (MainWindow)sender;
            if ((bool)e.NewValue)
            {
                syncWindow.txtWelcomeText.Text = "You are ready to start synchronizing your calendar. Click \"Synchronize\" to start the synchronization";
                syncWindow.imgOverlay.Source = new BitmapImage(new Uri("pack://application:,,,/Images/Apply64.png"));
            }
            else
            {
                syncWindow.txtWelcomeText.Text = "Before you can start synchronizing, you must enter your notes password and select the database which contains the calendar entries";
                syncWindow.imgOverlay.Source = new BitmapImage(new Uri("pack://application:,,,/Images/Fail64.png"));
            }
        }
        private NotesReader _notesReader;
        private OutlookManager _outlookManager;
        private CalendarMerger _calendarMerger;
        private ProgramSettings settings;
        
        public MainWindow()
        {
            InitializeComponent();
            progressBorder.Visibility = Visibility.Collapsed;
            settings = ProgramSettings.LoadSettings();
            txtNotesPassword.Password = settings.NotesPassword;            
            _notesReader = new NotesReader();
            _outlookManager = new OutlookManager();
            _calendarMerger = new CalendarMerger();
            wsReadNotes.SetupWorker(_notesReader.FetchCalendarWorker);
            wsReadOutlook.SetupWorker(_outlookManager.FetchCalendarWorker);
            wsMergeEntries.SetupWorker(_calendarMerger.Worker);
            wsApplyChanges.SetupWorker(_outlookManager.MergeCalendarWorker);
            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
        }

        void ResetWorkSteps()
        {
            wsReadNotes.Reset();
            wsReadOutlook.Reset();
            wsMergeEntries.Reset();
            wsApplyChanges.Reset();
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (settings.NotesPassword != null && settings.NotesPassword.Length > 0)
                RefreshNotesDatabases();
            else
                expSettings.IsExpanded = true;
            DryRun = settings.DryRun;
            UpdateIsReadyState();
            //new MergeWindow(new List<ModifiedEntry>()).Show();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            settings.DryRun = DryRun;
            settings.NotesDatabase = (string)cmbNotesDB.SelectedItem;
            settings.NotesPassword = txtNotesPassword.Password;
            settings.Save();
            base.OnClosing(e);
        }

        private void UpdateIsReadyState()
        {
            if (String.IsNullOrEmpty(settings.NotesDatabase) || String.IsNullOrEmpty(settings.NotesPassword))
                IsReadyToSynchronize = false;
            else
                IsReadyToSynchronize = true;
        }

        private void BeginFetchCalendarEntries()
        {
            IsSynchronizing = true;
            ResetWorkSteps();
            wsReadNotes.StartWork();
            wsReadOutlook.StartWork();
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
            UpdateIsReadyState();
        }

        private void btnRefreshNotesDB_Click(object sender, RoutedEventArgs e)
        {
            RefreshNotesDatabases();
        }

        private void cmbNotesDB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            settings.NotesDatabase = (string) cmbNotesDB.SelectedItem;
            UpdateIsReadyState();
        }
        private void DisplaySynchronizationStatus(string title, string message, InfoBoxType type)
        {
            progressInfoBox.Title = title;
            progressInfoBox.Message = message;
            progressInfoBox.InfoBoxType = type;
            progressInfoBox.ShowAndAutoClose();
        }

        private void DisplaySynchronizationStatus(WorkerStep ws)
        {
            switch (ws.WorkStage)
            {
                case WorkStepStage.Cancelled:
                    DisplaySynchronizationStatus("Operation cancelled", "The synchronization was cancelled. No changes was written to Outlook", InfoBoxType.Warning);
                    break;
                case WorkStepStage.Failed:
                    DisplaySynchronizationStatus("Error Reading Calendars", ws.ErrorMessage, InfoBoxType.Error);
                    break;
                default:
                    return;
            }
        }

        private void wsReadCalendar_WorkDone(object sender, RoutedEventArgs e)
        {
            if (wsReadOutlook.IsFinished && wsReadNotes.IsFinished)
            {
                if (wsReadOutlook.WorkStage == WorkStepStage.Completed && wsReadNotes.WorkStage == WorkStepStage.Completed)
                {
                    _calendarMerger.NotesEntries = _notesReader.CalendarEntries;
                    _calendarMerger.OutlookEntries = _outlookManager.CalendarEntries;
                    wsMergeEntries.StartWork();
                }
                else
                {
                    if (wsReadNotes.WorkStage != WorkStepStage.Completed)
                        DisplaySynchronizationStatus(wsReadNotes);
                    else
                        DisplaySynchronizationStatus(wsReadOutlook);
                }
            }
        }

        private void wsMergeEntries_WorkDone(object sender, RoutedEventArgs e)
        {
            if (wsMergeEntries.WorkStage == WorkStepStage.Completed)
            {
                if (_calendarMerger.ModifiedEntries.Count == 0)
                {
                    MessageBox.Show("Your calendars are already fully synchronized!");
                    // Call the final workstage anyway to limit the amount of extra code paths (the workstage will complete instantly anyway)
                    wsApplyChanges.StartWork(_calendarMerger.ModifiedEntries);
                    return;
                }
                MergeWindow mergeWin = new MergeWindow(_calendarMerger.ModifiedEntries);
                bool doMerge = (mergeWin.ShowDialog() == true);
                if (doMerge && !DryRun)
                {
                    wsApplyChanges.StartWork(_calendarMerger.ModifiedEntries);
                }
                else
                {
                    DisplaySynchronizationStatus("Merge cancelled", "The merge dialog was cancelled and thus no changes was written to Outlook", InfoBoxType.Warning);
                }
            }
            else
            {
                // Cancelled or failed
                DisplaySynchronizationStatus(wsMergeEntries);
            }
        }

        private void wsApplyChanges_WorkDone(object sender, RoutedEventArgs e)
        {
            if (wsApplyChanges.WorkStage == WorkStepStage.Completed)
            {
                _calendarMerger.SaveMappings();
                progressInfoBox.InfoBoxType = InfoBoxType.Info;
                progressInfoBox.Title = "Synchronization completed";
                var sb = new StringBuilder();
                sb.AppendFormat("A total of {0} (of {1} available modifications) was successfully merged with the Outlook calendar", _outlookManager.NumberOfMergedEntries, _calendarMerger.ModifiedEntries.Count); 
                if (_notesReader.NumberOfSkippedEntries > 0)
                    sb.AppendFormat("{0} calendar entries from Lotus Notes was skipped.", _notesReader.NumberOfSkippedEntries);
                if (_outlookManager.NumberOfSkippedEntries > 0)
                    sb.AppendFormat("{0} calendar entries from Outlook was ignored.", _outlookManager.NumberOfSkippedEntries);
                
                sb.AppendLine();
                sb.Append("You can now use iTunes to synchronize the Outlook calendar with your iPhone.");
                progressInfoBox.Message = sb.ToString();
                progressInfoBox.ShowAndAutoClose();
            }
            else
                // Cancelled or failed
                DisplaySynchronizationStatus(wsApplyChanges);
        }

        private void progressInfoBox_MessageConfirmed(object sender, RoutedEventArgs e)
        {
            IsSynchronizing = false;
        }

        private void btnQuit_Click(object sender, RoutedEventArgs e)
        {
            Close();
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
