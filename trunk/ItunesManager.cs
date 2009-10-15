using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTunesLib;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

namespace TieCal
{
    class ItunesManager
    {
        public ItunesManager()
        {
            SynchronizeWorker = new BackgroundWorker();
            SynchronizeWorker.DoWork += new DoWorkEventHandler(SynchronizeWorker_DoWork);
            SynchronizeWorker.WorkerSupportsCancellation = false;
            SynchronizeWorker.WorkerReportsProgress = true;
        }

        private IITIPodSource GetIpodById(iTunesApp app, ItunesId IphoneId)
        {
            foreach (IITSource src in app.Sources)
            {
                if (src.Kind == ITSourceKind.ITSourceKindIPod)
                {
                    object o = src;
                    ItunesId id = new ItunesId(app.get_ITObjectPersistentIDHigh(ref o), app.get_ITObjectPersistentIDLow(ref o));
                    if (id == IphoneId)
                        return src as IITIPodSource;
                }
            }
            return null;
        }
        /// <summary>
        /// Gets the first connected ipod in iTunes.
        /// </summary>
        /// <param name="app">The itunes app object.</param>
        /// <returns>An IPodSource or null if no ipod is connected</returns>
        private IITIPodSource GetFirstIpod(iTunesApp app)
        {
            foreach (IITSource src in app.Sources)
            {
                if (src.Kind == ITSourceKind.ITSourceKindIPod)
                    return src as IITIPodSource;
            }
            return null;
        }

        void SynchronizeWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                iTunesApp app;
                BackgroundWorker worker = (BackgroundWorker)sender;
                worker.ReportProgress(15);
                var start = Environment.TickCount;
                app = new iTunesAppClass();
                var time = Environment.TickCount - start;
                bool wasRunning = false;
                if (time < 500)
                    wasRunning = true;
                IITIPodSource iphone = null;
                worker.ReportProgress(20);                
                if (ProgramSettings.Instance.IphoneId.IsEmpty)
                {
                    iphone = GetFirstIpod(app);                    
                }
                else
                    iphone = GetIpodById(app, ProgramSettings.Instance.IphoneId);
                worker.ReportProgress(40);
                if (iphone == null)
                {
                    // no iphone connected
                    worker.ReportProgress(100);
                    e.Result = new ApplicationException("Failed to find a connected iPhone");
                    return;
                }
                object o = iphone;
                ProgramSettings.Instance.IphoneId = new ItunesId(app.get_ITObjectPersistentIDHigh(ref o), app.get_ITObjectPersistentIDLow(ref o));
                iphone.UpdateIPod();
                // Just fake progress.. this is bound to be inaccurate though :-(
                for (int i = 50; i < 100; i++)
                {
                    worker.ReportProgress(i);
                    Thread.Sleep(100);
                }
                // TODO: Figure out how to know when sync is completed so that we can close itunes and update our GUI
                //if (!wasRunning)
                //    app.Quit();
            }
            catch (COMException ex)
            {
                e.Result = new ApplicationException("Failed to communicate with iTunes: " + ex.Message, ex);
            }
        }       

        public BackgroundWorker SynchronizeWorker { get; private set; }
    }

    /// <summary>
    /// Represents an object in iTunes. These objects have 64bit persistent ids which are divided into two parts.
    /// </summary>
    public struct ItunesId
    {
        private int _high, _low;
        public Int32 High { get { return _high; } }
        public Int32 Low { get { return _low; } }
        public ItunesId(Int32 high, Int32 low) 
        {
            _high = high;
            _low = low;
        }

        public bool IsEmpty
        {
            get { return _high == 0 && _low == 0; }
        }

        public static bool operator ==(ItunesId left, ItunesId right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(ItunesId left, ItunesId right)
        {
            return !(left == right);
        }

        public override int GetHashCode()
        {
            return _low.GetHashCode() ^ _high.GetHashCode();
        }

        public bool Equals(ItunesId obj)
        {
            return obj._low == _low && obj._high == _high;
        }
        public override bool Equals(object obj)
        {
            return obj is ItunesId && Equals((ItunesId)obj);
        }
    }
}
