// Part of the TieCal project (http://code.google.com/p/tiecal/)
// Copyright (C) 2009, Isak Savo <isak.savo@gmail.com>
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//      http://www.gnu.org/licenses/gpl.html
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
	public partial class SetupRemindersDialog
	{
		public SetupRemindersDialog()
		{
			this.InitializeComponent();			
			// Insert code required on object creation below this point.
            if (System.ComponentModel.DesignerProperties.GetIsInDesignMode(this) == false)
            {
                ReminderMode = ProgramSettings.Instance.ReminderMode;
                if (ReminderMode == ReminderMode.Custom)
                    ReminderMinutes = ProgramSettings.Instance.ReminderMinutesBeforeStart;
                else
                    ReminderMinutes = 15;
            }
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

        public ReminderMode ReminderMode
        {
            get
            {
                if (rdoOutlook.IsChecked == true)
                    return ReminderMode.OutlookDefault;
                else if (rdoCustom.IsChecked == true)
                    return ReminderMode.Custom;
                else if (rdoDisable.IsChecked == true)
                    return ReminderMode.NoReminder;
                else
                    throw new InvalidOperationException("No reminder has been selected");
            }
            set
            {
                switch (value)
                {
                    case ReminderMode.NoReminder:
                        rdoDisable.IsChecked = true;
                        break;
                    case ReminderMode.OutlookDefault:
                        rdoOutlook.IsChecked = true;
                        break;
                    case ReminderMode.Custom:
                        rdoCustom.IsChecked = true;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("Reminder mode " + value + " is not known");
                }
            }
        }

        public int ReminderMinutes
        {
            // TODO: error checking
            get { return Int32.Parse(txtMinutes.Text); }
            set { txtMinutes.Text = value.ToString(); }
        }
	}
}