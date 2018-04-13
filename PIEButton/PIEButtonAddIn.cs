using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Windows.Forms;

namespace PIEButton
{
    public partial class PIEButtonAddIn
    {
        Outlook.Inspector inspector;
        Outlook.Explorer explorer;

        // Method for sending a phishing report while viewing target suspicous message in Outlook

        private void SendPhishingReport(Outlook.Inspector Inspector)
        {
            try
            {
                // Proceed with attempting to send the submission

                Outlook.MailItem phish_mail = Inspector.CurrentItem as Outlook.MailItem;
                Outlook.MailItem mail = Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                DateTime phish_date = DateTime.Now;
                mail.Subject = "New Phishing Report " + phish_date.ToString();
                Outlook.AddressEntry currentUser = Application.Session.CurrentUser.AddressEntry;

                // Target e-mail address (which will receive the phishing report) is defined in the associated config file variable, "ReportTargetEmail"
                //
                // You can also replace all instances of "Properties.Settings.Default.ReportTargetEmail" with a hard-coded string literal (or string variable name) if you'd prefer to not use the config file.
                // Example: mail.Recipients.Add("phishing.report@mydomain.com");

                mail.Recipients.Add(Properties.Settings.Default.ReportTargetEmail);

                mail.Recipients.ResolveAll();
                mail.Attachments.Add(phish_mail, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                mail.Send();
                Debug.WriteLine(String.Format("{0} PIE Button: Phishing report submitted", DateTime.Now));

                // Notify the end-user that the submission was sent

                DialogResult ds = MessageBox.Show(String.Format("Phishing report successfully submitted to: {0}\n\nThank you!", Properties.Settings.Default.ReportTargetEmail), "Report Phishing - Submission Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Close the suspicous message in Outlook explorer

                phish_mail.Close(Outlook.OlInspectorClose.olDiscard);
                Debug.WriteLine(String.Format("{0} PIEButton: Closed suspicious message", DateTime.Now));

                return;
            }
            catch (Exception ex)
            {
                // An error has occurred

                Debug.Write(String.Format("PIEButton: {0}", ex.ToString()));
                return;
            }
        }

        // Method for sending a phishing report while suspicious message is selected in Outlook's message explorer

        private void SendPhishingReportExplorer(Outlook.Explorer Explorer, Outlook.Inspector Inspector)
        {
            try
            {
                // Test whether the currently selected item in Outlook explorer is an e-mail message or not

                Object selObject = Explorer.Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    // Selected item is indeed an e-mail message, proceed with submission

                    Debug.Write("PIEButton: Selected object IS a mail item");
                    Outlook.MailItem phish_mail = Explorer.Selection[1] as Outlook.MailItem;
                    Outlook.MailItem mail = Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    DateTime phish_date = DateTime.Now;
                    mail.Subject = "New Phishing Report " + phish_date.ToString();
                    Outlook.AddressEntry currentUser = Application.Session.CurrentUser.AddressEntry;

                    // Target e-mail address (which will receive the phishing report) is defined in the associated config file variable, "ReportTargetEmail"
                    //
                    // You can also replace all instances of "Properties.Settings.Default.ReportTargetEmail" with a hard-coded string literal (or string variable name) if you'd prefer to not use the config file.
                    // Example: mail.Recipients.Add("phishing.report@mydomain.com");

                    mail.Recipients.Add(Properties.Settings.Default.ReportTargetEmail);

                    mail.Recipients.ResolveAll();
                    mail.Attachments.Add(phish_mail, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    mail.Send();
                    Debug.WriteLine(String.Format("{0} PIE Button: Phishing report submitted", DateTime.Now));

                    DialogResult ds = MessageBox.Show(String.Format("Phishing report successfully submitted to: {0}\n\nThank you!", Properties.Settings.Default.ReportTargetEmail), "Report Phishing - Submission Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    return;
                }
                else
                {
                    // Selected item is NOT an e-mail message, but no error has occurred. Do nothing.

                    Debug.Write("PIEButton: Selected object is not a mail item");

                    return;
                }
            }
            catch (Exception exp)
            {
                // No item is currently selected in Outlook explorer, or something has gone wrong. Display popup box notifying the user.

                DialogResult ds2 = MessageBox.Show("No e-mail message currently selected;\nPlease select a single message and attempt submission again.", "Report Phishing - No Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.Write(String.Format("PIEButton: {0}", exp.ToString()));

                return;
            }
        }

        public void MySendPhish()
        {
            ThisAddIn_SendPhishReport();
        }

        public void MySendPhishExplorer()
        {
            ThisAddIn_SendPhishReportExplorer();
        }

        public delegate void SendPhishingDelegate();

        private void ThisAddIn_SendPhishReport()
        {
            inspector = this.Application.ActiveInspector();
            SendPhishingReport(inspector);
        }

        private void ThisAddIn_SendPhishReportExplorer()
        {
            inspector = this.Application.ActiveInspector();
            explorer = this.Application.ActiveExplorer();
            SendPhishingReportExplorer(explorer, inspector);
        }

        public class PhishClass
        {
            public event SendPhishingDelegate SendPhish;
            public void DoPhishSend()
            {
                SendPhish();
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspector = this.Application.ActiveInspector();
            explorer = this.Application.ActiveExplorer();
            Debug.WriteLine(String.Format("{0} PIEButton: Starting up...", DateTime.Now));
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
