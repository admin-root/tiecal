using System;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.ComponentModel;

namespace TieCal
{
	public partial class InfoBox
    {
        #region Dependency Properties
        // Using a DependencyProperty as the backing store for Title.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register("Title", typeof(string), typeof(InfoBox), new UIPropertyMetadata(null));
        // Using a DependencyProperty as the backing store for Message.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty MessageProperty =
            DependencyProperty.Register("Message", typeof(string), typeof(InfoBox), new UIPropertyMetadata(""));
        // Using a DependencyProperty as the backing store for Icon.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IconProperty =
            DependencyProperty.Register("Icon", typeof(ImageSource), typeof(InfoBox), new UIPropertyMetadata(null));

        // Using a DependencyProperty as the backing store for InfoBoxType.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty InfoBoxTypeProperty =
            DependencyProperty.Register("InfoBoxType", typeof(InfoBoxType), typeof(InfoBox), new UIPropertyMetadata(InfoBoxType.Info));

        // Using a DependencyProperty as the backing store for IconSize.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IconSizeProperty =
            DependencyProperty.Register("IconSize", typeof(double), typeof(InfoBox), new UIPropertyMetadata(48.0));

        // Using a DependencyProperty as the backing store for ShowConfirmButton.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ShowConfirmButtonProperty =
            DependencyProperty.Register("ShowConfirmButton", typeof(bool), typeof(InfoBox), new UIPropertyMetadata(false));

        /// <summary>
        /// Gets or sets a value indicating whether to show the confirmation button. This is a dependency property
        /// </summary>
        [Description("Gets or sets a value indicating whether to show the confirmation button."), Category("Common Properties")]
        public bool ShowConfirmButton
        {
            get { return (bool)GetValue(ShowConfirmButtonProperty); }
            set { SetValue(ShowConfirmButtonProperty, value); }
        }

        /// <summary>
        /// Gets or sets the size of the icon, in device independent pixels. This is a dependency property.
        /// </summary>
        /// <value>The size of the icon.</value>
        [Description("Gets or sets the size of the icon, in device independent pixels."), Category("Common Properties")]
        public double IconSize
        {
            get { return (double)GetValue(IconSizeProperty); }
            set { SetValue(IconSizeProperty, value); }
        }


        /// <summary>
        /// Gets or sets the type of the info box. This is a dependency property.
        /// </summary>
        [Description("Gets or sets the type of the info box."), Category("Appearance")]
        public InfoBoxType InfoBoxType
        {
            get { return (InfoBoxType)GetValue(InfoBoxTypeProperty); }
            set { SetValue(InfoBoxTypeProperty, value); }
        }

        /// <summary>
        /// Gets or sets the title of the message. This is a dependency property.
        /// </summary>
        /// <value>The title.</value>
        [Description("Gets or sets the title of the message."), Category("Common Properties")]
        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }

        /// <summary>
        /// Gets or sets the message to display. This is a dependency property.
        /// </summary>
        [Description("Gets or sets the message to display."), Category("Common Properties")]
        public string Message
        {
            get { return (string)GetValue(MessageProperty); }
            set { SetValue(MessageProperty, value); }
        }

        /// <summary>
        /// Gets or sets the icon to display next to the message. This is a dependency property.
        /// </summary>
        /// <value>The icon.</value>
        [Description("Gets or sets the icon to display next to the message."), Category("Common Properties")]
        public ImageSource Icon
        {
            get { return (ImageSource)GetValue(IconProperty); }
            set { SetValue(IconProperty, value); }
        }
        #endregion

        public static RoutedEvent MessageConfirmedEvent = EventManager.RegisterRoutedEvent("MessageConfirmed", RoutingStrategy.Bubble, 
                                                            typeof(RoutedEventHandler), typeof(InfoBox));
        /// <summary>
        /// Raised when the user has acknowledged the message
        /// </summary>
        public event RoutedEventHandler MessageConfirmed
        {
            add { AddHandler(MessageConfirmedEvent, value); }
            remove { RemoveHandler(MessageConfirmedEvent, value); }
        }

        public InfoBox()
		{
			this.InitializeComponent();            
		}

        private bool AutoClose { get; set; }

        /// <summary>
        /// Shows the infobox and then when the user clicks the confirmation button, the infobox closes itself (by setting Visibility to Collapsed).
        /// </summary>
        public void ShowAndAutoClose()
        {
            AutoClose = true;
            ShowConfirmButton = true;
            this.Visibility = Visibility.Visible;
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            if (AutoClose)
                this.Visibility = Visibility.Collapsed;
            AutoClose = false;
            RaiseEvent(new RoutedEventArgs(MessageConfirmedEvent));
            
        }
	}

    /// <summary>
    /// Represents the different styles an <see cref="InfoBox"/> can be
    /// </summary>
    public enum InfoBoxType
    {
        /// <summary>Uses an error icon with red background</summary>
        Error,
        /// <summary>Uses a warning icon with yellow/orange background</summary>        
        Warning,
        /// <summary>Uses an information icon with blue background</summary>
        Info,
        /// <summary>Don't use any built in logic to set icon and background</summary>
        Custom,
    }

    /// <summary>
    /// Converter that allows an element to be collapsed if the value it is bound to is <c>null</c>
    /// </summary>
    public class NullToVisibilityConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
                return Visibility.Collapsed;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

}