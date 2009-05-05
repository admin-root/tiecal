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
        void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            try
            {
                StringBuilder sb = new StringBuilder("Unhandled Exception: ");
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
                MessageBox.Show(sb.ToString());

            }
            catch
            {
                /* We can't allow a throw in this method */
            }
            finally
            {
                e.Handled = true;
            }
        }
    }
}
