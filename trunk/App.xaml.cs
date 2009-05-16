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
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.Text;

namespace TieCal
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += new System.Windows.Threading.DispatcherUnhandledExceptionEventHandler(App_DispatcherUnhandledException);
        }

        private string GetExceptionDetailsString(Exception ex, int indent)
        {
            StringBuilder sb = new StringBuilder();
            string padding = new string(' ', indent);
            sb.AppendFormat("{0}Type: {1}", padding, ex.GetType().FullName);
            sb.AppendLine();
            sb.AppendFormat("{0}Message: {1}", padding, ex.Message);
            sb.AppendLine();
            sb.AppendLine(padding + "Details:");
            sb.Append(padding + padding);
            if (ex is System.IO.FileNotFoundException)
                sb.AppendFormat("FileName: {0}", (ex as System.IO.FileNotFoundException).FileName);
            else if (ex is ArgumentException)
                sb.AppendFormat("Parameter: {0}", padding, (ex as ArgumentException).ParamName);
            return sb.ToString();
        }
        private bool IsComException(Exception ex)
        {
            while (ex != null)
            {
                if (ex is System.Runtime.InteropServices.COMException)
                    return true;
                ex = ex.InnerException;
            }
            return false;
        }
        void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            try
            {
                // Tell Windows we've handled the error, otherwise it'll try to submit a crash report to microsoft.com
                e.Handled = true;
                StringBuilder sb = new StringBuilder();
                if (IsComException(e.Exception))
                {
                    sb.AppendLine("Failed to communicate with calendar applications. This usually means that Outlook and/or Lotus Notes isn't installed.");
                    sb.AppendLine("");
                    sb.AppendLine("Tested versions are: Lotus Notes 7.0.2 and Outlook 2007");
                    sb.AppendLine("TieCal will now exit");
                    MessageBox.Show(sb.ToString());
                    Environment.Exit(-2);
                    return;
                }
                sb.AppendLine("Unhandled Exception: ");
                if (e.Exception != null)
                {
                    sb.AppendLine(GetExceptionDetailsString(e.Exception, 0));
                    Exception ex = e.Exception.InnerException;
                    int indent = 2;
                    while (ex != null)
                    {
                        sb.AppendLine(GetExceptionDetailsString(ex, indent));
                        ex = ex.InnerException;
                        indent += 2;
                    }
                    sb.AppendLine();
                    sb.AppendLine("StackTrace: " + e.Exception.StackTrace);
                }
                sb.AppendLine("");
                sb.AppendLine("Press 'Cancel' to terminate or 'Ok' to keep the application running");
                var response = MessageBox.Show(sb.ToString(), "TieCal Error", MessageBoxButton.OKCancel);
                if (response != MessageBoxResult.OK)
                    Environment.Exit(-1);
            }
            catch
            {
                /* We can't allow a throw in this method */
            }
        }
    }
}
