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
    /// Interaction logic for SyncConfirmation.xaml
    /// </summary>
    public partial class SyncConfirmation : Window
    {
        public bool IsDontAskAgain
        {
            get
            {
                return chkDontAskAgain.IsChecked.HasValue &&
                    chkDontAskAgain.IsChecked.Value;
            }
        }

        public SyncConfirmation()
        {
            InitializeComponent();
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
