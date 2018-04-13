using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;

namespace PIEButton
{
    public partial class PIERibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Debug.WriteLine(String.Format("{0} PIEButton: Ribbon loading...", DateTime.Now));
        }

        private void pieButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Debug.WriteLine(String.Format("{0} PIEButton: Reporting phishing from Outlook explorer...", DateTime.Now));
            Globals.PIEButtonAddIn.MySendPhishExplorer();
        }

        private void pieButton2_Click(object sender, RibbonControlEventArgs e)
        {
            Debug.WriteLine(String.Format("{0} PIEButton: Reporting phishing message...", DateTime.Now));
            Globals.PIEButtonAddIn.MySendPhish();
        }

    }
}
