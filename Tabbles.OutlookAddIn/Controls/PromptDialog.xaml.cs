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

namespace Tabbles.OutlookAddIn.Controls
{
    /// <summary>
    /// Interaction logic for PromptDialog.xaml
    /// </summary>
    public partial class PromptDialog : Window
    {
        public string Message
        {
            get;
            set;
        }

        private string okText;
        public string OkText
        {
            get
            {
                return okText;
            }
            set
            {
                okText = value;
            }
        }

        public string CancelText
        {
            get;
            set;
        }

        public string DontShowAgainMessage
        {
            get;
            set;
        }

        public bool IsDontAskAgain
        {
            get
            {
                return chkDontAskAgain.IsChecked.HasValue &&
                    chkDontAskAgain.IsChecked.Value;
            }
        }

        public bool WasDontAskAgain
        {
            get;
            set;
        }

        public PromptDialog()
        {
            this.Loaded += PromptDialog_Loaded;

            InitializeComponent();
        }

        private void PromptDialog_Loaded(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Message))
            {
                txbMessage.Visibility = Visibility.Visible;
                txbMessage.Text = Message;
            }

            if (!string.IsNullOrEmpty(OkText))
            {
                btnOk.Content = OkText;
            }

            if (!string.IsNullOrEmpty(CancelText))
            {
                btnCancel.Content = CancelText;
            }

            if (!string.IsNullOrEmpty(DontShowAgainMessage))
            {
                chkDontAskAgain.Visibility = Visibility.Visible;
                chkDontAskAgain.Content = DontShowAgainMessage;
                chkDontAskAgain.IsChecked = WasDontAskAgain;
            }
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void chkDontAskAgain_Checked(object sender, RoutedEventArgs e)
        {
            btnOk.IsEnabled = false;
        }

        private void chkDontAskAgain_Unchecked(object sender, RoutedEventArgs e)
        {
            btnOk.IsEnabled = true;
        }
    }
}
