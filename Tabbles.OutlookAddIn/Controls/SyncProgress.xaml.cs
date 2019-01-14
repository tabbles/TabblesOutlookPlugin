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
using System.Threading;
using Res = Tabbles.OutlookAddIn.Properties.Resources;
using System.Runtime.Remoting.Messaging;

namespace Tabbles.OutlookAddIn.Controls
{
    /// <summary>
    /// Interaction logic for SyncProgress.xaml
    /// </summary>
    public partial class SyncProgress : Window
    {
        private Action actionToRun;

        //public event EventHandler Cancel;

        public SyncProgress(Action actionToRun)
        {
            this.actionToRun = actionToRun;

            this.Closing += SyncProgress_Closing;

            InitializeComponent();

            actionToRun.BeginInvoke(SyncFinished, null);

            ////close after successful sync
            //this.syncThread.Join();
            //DialogResult = true;
        }

        private void SyncFinished(IAsyncResult asyncResult)
        {
            Action syncAction = (Action)((AsyncResult)asyncResult).AsyncDelegate;
            try
            {
                syncAction.EndInvoke(asyncResult);
            }
            catch (Exception)
            {
            }
            finally
            {
                Dispatcher.BeginInvoke(new ThreadStart(() =>
                {
                    DialogResult = true;
                }));
            }
        }

        private void SyncProgress_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (DialogResult.HasValue && DialogResult.Value)
            {
                return;
            }

            PromptDialog prompt = new PromptDialog()
            {
                Message = Res.MsgAreYouSure,
                OkText = Res.LabelYes,
                CancelText = Res.LabelNo
            };

            bool? answer = prompt.ShowDialog();
            if (!answer.HasValue || !answer.Value)
            {
                e.Cancel = true;
            }
            else
            {
            //    if (Cancel != null)
            //    {
            //        Cancel(this, EventArgs.Empty);
            //    }
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
