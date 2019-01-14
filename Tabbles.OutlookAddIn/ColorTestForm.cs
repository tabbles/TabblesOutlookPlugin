using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Tabbles.OutlookAddIn
{
    public partial class ColorTestForm : Form
    {
        public ColorTestForm()
        {
            InitializeComponent();
        }

        private void btnEval_Click(object sender, EventArgs e)
        {
            try
            {
                Outlook.OlCategoryColor color = Utils.GetOutlookColorFromRgb("#FF" + txtInput.Text.Trim());
                lblResult.Text = color.ToString();
                //lblResult.BackColor = Utils.OutlookColorsRgb[color];
            }
            catch (Exception)
            {
            }
        }
    }
}
