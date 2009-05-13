﻿using System;
using System.Collections.Generic;
using System.Windows;
using System.ComponentModel;
using System.Windows.Media.Imaging;
using System.Windows.Media;

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
        /// <summary>
        /// Gets or sets the title of this work step. This is a dependency property
        /// </summary>
        /// <value>The title.</value>
        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Title.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register("Title", typeof(string), typeof(WorkerStep), new UIPropertyMetadata("Working"));

        public WorkStepStage WorkStage
        {
            get { return (WorkStepStage)GetValue(WorkStageProperty); }
            private set { SetValue(WorkStageKey, value); }
        }

        private static readonly DependencyPropertyKey WorkStageKey =
            DependencyProperty.RegisterReadOnly("WorkStage", typeof(WorkStepStage), typeof(WorkerStep), new UIPropertyMetadata(WorkStepStage.Waiting, new PropertyChangedCallback(WorkStage_Changed)));

        public static readonly DependencyProperty WorkStageProperty = WorkStageKey.DependencyProperty;

        //public Brush BorderBrush
        //{
        //    get { return (Brush)GetValue(BorderBrushProperty); }
        //    set { SetValue(BorderBrushProperty, value); }
        //}

        //// Using a DependencyProperty as the backing store for BorderBrush.  This enables animation, styling, binding, etc...
        //public static readonly DependencyProperty BorderBrushProperty =
        //    DependencyProperty.Register("BorderBrush", typeof(Brush), typeof(WorkerStep), new UIPropertyMetadata(null));

        public Brush BorderBackground
        {
            get { return (Brush)GetValue(BorderBackgroundProperty); }
            set { SetValue(BorderBackgroundProperty, value); }
        }

        // Using a DependencyProperty as the backing store for BorderBackgroundBrush.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BorderBackgroundProperty =
            DependencyProperty.Register("BorderBackground", typeof(Brush), typeof(WorkerStep), new UIPropertyMetadata(null));

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
            worker.RunWorkerAsync(argument);
            WorkStage = WorkStepStage.Working;
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            pbar.Value = 100;
            if (e.Error != null)
            {
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
	}

    public enum WorkStepStage
    {
        Waiting,
        Working,
        Cancelled,
        Failed,
        Completed
    }
}