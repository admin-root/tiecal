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
	public partial class PasswordDialog
	{

        /// <summary>
        /// Gets or sets a value indicating whether the program should remember the password or not
        /// </summary>
        public bool RememberPassword
        {
            get { return (bool)GetValue(RememberPasswordProperty); }
            set { SetValue(RememberPasswordProperty, value); }
        }

        // Using a DependencyProperty as the backing store for RememberPassword.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty RememberPasswordProperty =
            DependencyProperty.Register("RememberPassword", typeof(bool), typeof(PasswordDialog), new UIPropertyMetadata(false));

		public PasswordDialog()
		{
			this.InitializeComponent();
			// Insert code required on object creation below this point.
            this.Loaded += new RoutedEventHandler(PasswordDialog_Loaded);
        }

        void PasswordDialog_Loaded(object sender, RoutedEventArgs e)
        {
            txtPassword.Focus();
        }
        
        /// <summary>
        /// Gets or sets the password in the dialog .
        /// </summary>
        public string Password 
        {
            get { return txtPassword.Password; }
            set { txtPassword.Password = value; } 
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