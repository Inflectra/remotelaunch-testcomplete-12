using System;
using System.Collections.Generic;
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

namespace Inflectra.RemoteLaunch.Engines.TestCompleteEngine.TestComplete12
{
    /// <summary>
    /// Interaction logic for EngineSettings.xaml
    /// </summary>
    public partial class EngineSettings : UserControl
    {
        public EngineSettings()
        {
            InitializeComponent();
            this.LoadSettings();
        }

        /// <summary>
        /// Loads the saved settings
        /// </summary>
        private void LoadSettings()
        {
            //Load the various properties
            this.txtWaitTime.Text = Properties.Settings.Default.WaitTimeMilliseconds.ToString();
            this.chkApplicationVisible.IsChecked = Properties.Settings.Default.ApplicationVisible;
            this.chkCloseAfterTest.IsChecked = Properties.Settings.Default.CloseApplicationAfterTest;
            this.chkIncludeEventMessages.IsChecked = Properties.Settings.Default.IncludeEventMessages;
        }

        /// <summary>
        /// Saves the specified settings.
        /// </summary>
        public void SaveSettings()
        {
            //Get the various properties
            int waitTime;
            if (Int32.TryParse(this.txtWaitTime.Text, out waitTime))
            {
                Properties.Settings.Default.WaitTimeMilliseconds = waitTime;
            }
            if (this.chkApplicationVisible.IsChecked.HasValue)
            {
                Properties.Settings.Default.ApplicationVisible = this.chkApplicationVisible.IsChecked.Value;
            }
            if (this.chkCloseAfterTest.IsChecked.HasValue)
            {
                Properties.Settings.Default.CloseApplicationAfterTest = this.chkCloseAfterTest.IsChecked.Value;
            }
            if (this.chkIncludeEventMessages.IsChecked.HasValue)
            {
                Properties.Settings.Default.IncludeEventMessages = this.chkIncludeEventMessages.IsChecked.Value;
            }

            //Save the properties and reload
            Properties.Settings.Default.Save();
            this.LoadSettings();
        }
    }
}
