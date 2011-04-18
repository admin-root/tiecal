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
using System.Windows;
using System.ComponentModel;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Windows.Data;
using System.IO;
using System.Diagnostics;

namespace TieCal
{
	public partial class WorkerStep
    {
        #region Dependency Properties & Routed Events

        public static readonly RoutedEvent WorkDoneEvent = EventManager.RegisterRoutedEvent("WorkDone", RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(WorkerStep));
        public event RoutedEventHandler WorkDone
        {
            add { AddHandler(WorkDoneEvent, value); }
            remove { RemoveHandler(WorkDoneEvent, value); }
        }
        
        public static readonly DependencyProperty IsAbortableProperty =
            DependencyProperty.Register("IsAbortable", typeof(bool), typeof(WorkerStep), new UIPropertyMetadata(false));

        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register("Title", typeof(string), typeof(WorkerStep), new UIPropertyMetadata("Working"));

        /// <summary>
        /// Gets or sets a value indicating whether this workerstep can be aborted. This is a dependency property
        /// </summary>
        public bool IsAbortable
        {
            get { return (bool)GetValue(IsAbortableProperty); }
            set { SetValue(IsAbortableProperty, value); }
        }

        /// <summary>
        /// Gets or sets the title of this work step. This is a dependency property
        /// </summary>
        /// <value>The title.</value>
        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }

        /// <summary>
        /// Gets the current work stage. This is a dependency property.
        /// </summary>
        public WorkStepStage WorkStage
        {
            get { return (WorkStepStage)GetValue(WorkStageProperty); }
            private set { SetValue(WorkStageKey, value); }
        }

        /// <summary>
        /// Gets or sets the border background. This is a dependency property.
        /// </summary>
        /// <value>The border background.</value>
        public Brush BorderBackground
        {
            get { return (Brush)GetValue(BorderBackgroundProperty); }
            set { SetValue(BorderBackgroundProperty, value); }
        }

        private static readonly DependencyPropertyKey WorkStageKey = DependencyProperty.RegisterReadOnly("WorkStage", typeof(WorkStepStage), typeof(WorkerStep), new UIPropertyMetadata(WorkStepStage.Waiting, new PropertyChangedCallback(WorkStage_Changed)));

        public static readonly DependencyProperty WorkStageProperty = WorkStageKey.DependencyProperty;

        // Using a DependencyProperty as the backing store for BorderBackgroundBrush.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BorderBackgroundProperty =
            DependencyProperty.Register("BorderBackground", typeof(Brush), typeof(WorkerStep), new UIPropertyMetadata(null));

        /// <summary>
        /// Gets or sets the status image to display in the box. This is a dependency property.
        /// </summary>
        /// <value>The status image.</value>
        public ImageSource StatusImage
        {
            get { return (ImageSource)GetValue(StatusImageProperty); }
            set { SetValue(StatusImageProperty, value); }
        }

        // Using a DependencyProperty as the backing store for StatusImage.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty StatusImageProperty =
            DependencyProperty.Register("StatusImage", typeof(ImageSource), typeof(WorkerStep), new UIPropertyMetadata(null));
        #endregion

        private BackgroundWorker worker = null;

        private static void WorkStage_Changed(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            WorkerStep ws = (WorkerStep)sender;
            switch ((WorkStepStage)e.NewValue)
            {
                case WorkStepStage.Failed:
                case WorkStepStage.Cancelled:
                case WorkStepStage.Completed:
                    ws.RaiseEvent(new RoutedEventArgs(WorkDoneEvent, ws));
                    break;
            }
        }

        public WorkerStep()
        {
            InitializeComponent();
        }

        public void Reset()
        {
            WorkStage = WorkStepStage.Waiting;
            ErrorMessage = null;
            pbar.Value = 0.0;
        }

        public void SetupWorker(BackgroundWorker worker)
		{
            if (this.worker != null)
            {
                this.worker.RunWorkerCompleted -= worker_RunWorkerCompleted;
                this.worker.ProgressChanged -= worker_ProgressChanged;
            }            
            worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
            this.worker = worker;
		}
        /// <summary>
        /// Gets a value indicating whether the worker has completed. Note that a worker that failed or was cancelled is also considered completed.
        /// </summary>
        public bool IsFinished
        {
            get { return WorkStage == WorkStepStage.Completed || WorkStage == WorkStepStage.Failed || WorkStage == WorkStepStage.Cancelled; }
        }
        private string _errorMessage;
        public string ErrorMessage
        {
            get
            {
                if (WorkStage != WorkStepStage.Failed)
                    return null;
                return _errorMessage;
            }
            set
            {
                _errorMessage = value;
            }
        }
        /// <summary>
        /// Starts the work that is associated with this workstep.
        /// </summary>
        public void StartWork()
        {
            StartWork(null);
        }

        /// <summary>
        /// Starts the work that is associated with this workstep.
        /// </summary>
        /// <param name="argument">The argument to pass to the background worker.</param>
        public void StartWork(object argument)
        {
            if (worker == null)
                throw new InvalidOperationException("No worker has been assigned to this WorkStep. Make sure you call SetupWorker before calling StartWork");
            this.IsAbortable = worker.WorkerSupportsCancellation;
            worker.RunWorkerAsync(argument);
            WorkStage = WorkStepStage.Working;
        }

        private void WriteSourceLineInfo(TextWriter writer, StackTrace trace)
        {
            var frame = trace.GetFrame(0);
            writer.WriteLine(" Source Location: {0}:{1},{2}", frame.GetFileName(), frame.GetFileLineNumber(), frame.GetFileColumnNumber());
        }

        private void LogException(Exception ex)
        {
            if (ex == null)
                return;
            string logfile = Path.Combine(ProgramSettings.SaveFolder, "crashlog.txt");            
            using (TextWriter writer = new StreamWriter(logfile, true))
            {
                writer.WriteLine("===================  {0} =======================", DateTime.Now.ToString());
                writer.WriteLine("Exception from {0}: {1}", Title, ex.Message);                
                WriteSourceLineInfo(writer, new StackTrace(ex, true)); 
                writer.WriteLine("  Stacktrace: {0}", ex.StackTrace.ToString());
                Exception innerEx = ex.InnerException;
                while (innerEx != null)
                {
                    writer.WriteLine("[inner exception] {0}:{1}", innerEx.GetType(), innerEx.Message);
                    WriteSourceLineInfo(writer, new StackTrace(ex, true));
                    writer.WriteLine(" StackTrace: {0}", innerEx.StackTrace);
                    innerEx = innerEx.InnerException;
                }
                writer.WriteLine("================================================");
            }
        }
        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            pbar.Value = 100;
            if (e.Error != null)
            {
                try
                {
                    LogException(e.Error);
                }
                catch { }
                ErrorMessage = e.Error.Message;
                WorkStage = WorkStepStage.Failed;
            }
            else if (e.Cancelled)
                WorkStage = WorkStepStage.Cancelled;
            else
                WorkStage = WorkStepStage.Completed;
        }


        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbar.Value = e.ProgressPercentage;
        }

        private void btnAbort_Click(object sender, RoutedEventArgs e)
        {
            if (worker.IsBusy)
                worker.CancelAsync();
        }
	}

    public enum WorkStepStage
    {
        Waiting,
        Working,
        Cancelled,
        Failed,
        Completed
    }

    public class WorkStepStageToEnabledConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            WorkStepStage stage = (WorkStepStage)value;
            if (stage == WorkStepStage.Working)
                return true;
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

}