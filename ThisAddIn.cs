using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Threading;

namespace Outbox_Monitor
{
    public partial class ThisAddIn
    {
        private enum LOGTYPE
        {
            LT_INFORMATION,
            LT_WARNING,
            LT_ERROR,
            LT_NONE
        };
        private bool running = false;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                WriteLog(LOGTYPE.LT_INFORMATION, "Outbox monitor started.");
                running = true;
                Thread backgroundThread = new Thread(new ThreadStart(CheckAndMoveSentItemsThread));
                backgroundThread.Start();
                WriteLog(LOGTYPE.LT_INFORMATION, "Outbox monitor background thread started.");
            }
            catch (Exception ex)
            {
                WriteLog(LOGTYPE.LT_ERROR, "Failed to start Outbox monitor: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private List<Outlook.MailItem> GetSentItems()
        {
            try
            {
                List<Outlook.MailItem> SentItems = new List<Outlook.MailItem>();
                Outlook.Folder outbox = (Outlook.Folder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
                Outlook.Items items = outbox.Items;
                if (items.Count > 0)
                {
                    string searchString = "[SentOn] < '" + DateTime.Now.AddSeconds(-60).ToShortDateString() + " " + DateTime.Now.AddSeconds(-60).ToShortTimeString() + "'";
                    object SentItem = items.Find(searchString);
                    while (SentItem != null)
                    {
                        SentItems.Add(SentItem as Outlook.MailItem);
                        SentItem = items.FindNext();
                    }
                }
                return SentItems;
            }
            catch (Exception ex)
            {
                WriteLog(LOGTYPE.LT_ERROR, ex.Message);
                return null;
            }
        }

        public void ListSentItems()
        {
            List<Outlook.MailItem> SentItems = GetSentItems();
            if (SentItems != null && SentItems.Count > 0)
            {
                String StrSentItems = "There " + ((SentItems.Count > 1) ? "are" : "is") + " " + SentItems.Count + " item" + ((SentItems.Count > 1) ? "s" : "") + " are in the outbox that have been sent:\r\n";
                foreach (Outlook.MailItem SentItem in SentItems)
                {
                    StrSentItems += SentItem.Subject + " - " + SentItem.SentOn.ToString() + "\r\n";
                }
                WriteLog(LOGTYPE.LT_INFORMATION, StrSentItems);
            }
        }

        public void CheckAndMoveSentItems()
        {
            List<Outlook.MailItem> SentItems = GetSentItems();
            if (SentItems != null && SentItems.Count > 0)
            {
                String StrSentItems = "Moved " + SentItems.Count + " item" + ((SentItems.Count > 1) ? "s" : "") + " to sent items:\r\n";
                foreach (Outlook.MailItem SentItem in SentItems)
                {
                    SentItem.Move((Outlook.Folder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail));
                    StrSentItems += SentItem.Subject + " - " + SentItem.SentOn.ToString() + "\r\n";
                }
                WriteLog(LOGTYPE.LT_INFORMATION, StrSentItems);
            }
        }

        private void CheckAndMoveSentItemsThread()
        {
            while(running == true)
            {
                CheckAndMoveSentItems();
                Thread.Sleep(60 * 1000);
            }
        }

        private void WriteLog(LOGTYPE LogType, string LogString)
        {
            try
            {
                CheckLogPaths();
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ": " + LogString);
                string LogPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Outbox Monitor\\Logs";
                if (LogType == LOGTYPE.LT_ERROR)
                {
                    LogPath += "\\Error.log";
                }
                else if (LogType == LOGTYPE.LT_WARNING)
                {
                    LogPath += "\\Warning.log";
                }
                else if (LogType == LOGTYPE.LT_INFORMATION)
                {
                    LogPath += "\\Information.log";
                }
                else
                {
                    return;
                }
                var LogFile = File.AppendText(LogPath);
                LogFile.WriteLine(sb);
                LogFile.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(null, ex.Message, "Error Writing to Log File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CheckLogPaths()
        {
            try
            {
                string LocalAppdataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (Directory.Exists(LocalAppdataPath + "\\Outbox Monitor") == false)
                {
                    Directory.CreateDirectory(LocalAppdataPath + "\\Outbox Monitor");
                }
                if (Directory.Exists(LocalAppdataPath + "\\Outbox Monitor\\Logs") == false)
                {
                    Directory.CreateDirectory(LocalAppdataPath + "\\Outbox Monitor\\Logs");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(null, ex.Message, "Error Creating Log Folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
