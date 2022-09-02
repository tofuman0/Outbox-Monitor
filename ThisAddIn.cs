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
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Outbox_Monitor
{
    public partial class ThisAddIn
    {
        private enum LOGTYPE
        {
            LT_ERROR,
            LT_WARNING,
            LT_INFORMATION,
            LT_NONE
        };
        private class Config
        {
            public bool? BackgroundMonitor { get; set; }
            public Int32? BackgroundInterval { get; set; }
            public LOGTYPE? LogLevel { get; set; }
            public bool? LogOnly { get; set; }
        }
        private Int32 lastHash = 0;
        Config config = new Config
        {
            BackgroundMonitor = true,
            BackgroundInterval = 60,
            LogLevel = LOGTYPE.LT_INFORMATION,
            LogOnly = false
        };
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                LoadConfig();
                WriteLog(LOGTYPE.LT_INFORMATION, "Outbox monitor started.");
                if (config.BackgroundMonitor == true)
                {
                    Thread backgroundThread = new Thread(new ThreadStart(CheckAndMoveSentItemsThread));
                    backgroundThread.Start();
                    WriteLog(LOGTYPE.LT_INFORMATION, "Outbox monitor background thread started.");
                }
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

        private List<object> GetSentItems()
        {
            try
            {
                List<object> SentItems = new List<object>();
                Outlook.Folder outbox = (Outlook.Folder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
                Outlook.Items items = outbox.Items;

                if (items.Count > 0)
                {
                    string searchString = "[SentOn] < '" + DateTime.Now.AddSeconds(-60).ToShortDateString() + " " + DateTime.Now.AddSeconds(-60).ToShortTimeString() + "'";
                    object SentItem = items.Find(searchString);
                    while (SentItem != null)
                    {
                        SentItems.Add(SentItem);
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

        public void GetOutboxItemTypes()
        {
            try
            {
                List<object> SentItems = new List<object>();
                Outlook.Folder outbox = (Outlook.Folder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
                Outlook.Items items = outbox.Items;

                foreach (object item in items)
                {
                    if(item is Outlook.MailItem)
                        WriteLog(LOGTYPE.LT_INFORMATION, "Item Type: MailItem");
                    else if (item is Outlook.AppointmentItem)
                        WriteLog(LOGTYPE.LT_INFORMATION, "Item Type: AppointmentItem");
                    else if (item is Outlook.MeetingItem)
                        WriteLog(LOGTYPE.LT_INFORMATION, "Item Type: MeetingItem");
                    else
                        WriteLog(LOGTYPE.LT_INFORMATION, "Item Type: Unknown");
                }
            }
            catch (Exception ex)
            {
                WriteLog(LOGTYPE.LT_ERROR, ex.Message);
            }
        }

        public void LogSentItems()
        {
            try
            {
                List<object> SentItems = GetSentItems();
                // Prevent multiple logs while just logging the items in the outbox
                if (lastHash != GetHash(SentItems))
                {
                    if (SentItems != null && SentItems.Count > 0)
                    {
                        String StrSentItems = "There " + ((SentItems.Count > 1) ? "are" : "is") + " " + SentItems.Count + " item" + ((SentItems.Count > 1) ? "s" : "") + " are in the outbox that have been sent:\r\n";
                        for (Int32 i = 0; i < SentItems.Count; i++)
                        {
                            if (SentItems[i] == null)
                            {
                                WriteLog(LOGTYPE.LT_ERROR, "Outbox item at index " + i + " is null.");
                            }
                            else
                            {
                                if (SentItems[i] is Outlook.MailItem)
                                {
                                    Outlook.MailItem SentItem = (Outlook.MailItem)SentItems[i];
                                    StrSentItems += SentItem.Subject + " - " + SentItem.SentOn.ToString() + "\r\n";
                                }
                                else if (SentItems[i] is Outlook.MeetingItem)
                                {
                                    Outlook.MeetingItem SentItem = (Outlook.MeetingItem)SentItems[i];
                                    StrSentItems += SentItem.Subject + " - " + SentItem.SentOn.ToString() + "\r\n";
                                }
                                else
                                    continue;
                            }
                        }
                        WriteLog(LOGTYPE.LT_INFORMATION, StrSentItems);
                    }
                    lastHash = GetHash(SentItems);
                }
            }
            catch (Exception ex)
            {
                WriteLog(LOGTYPE.LT_ERROR, ex.Message);
            }
        }

        public void CheckAndMoveSentItems()
        {
            try
            {
                if (config.LogOnly.HasValue && config.LogOnly == true)
                {
                    LogSentItems();
                    return;
                }
                List<object> SentItems = GetSentItems();
                if (SentItems != null && SentItems.Count > 0)
                {
                    String StrSentItems = "Moved " + SentItems.Count + " item" + ((SentItems.Count > 1) ? "s" : "") + " to sent items:\r\n";
                    for (Int32 i = 0; i < SentItems.Count; i++)
                    {
                        if (SentItems[i] == null)
                        {
                            WriteLog(LOGTYPE.LT_ERROR, "Outbox item at index " + i + " is null.");
                        }
                        else
                        {
                            if (SentItems[i] is Outlook.MailItem)
                            {
                                Outlook.MailItem SentItem = (Outlook.MailItem)SentItems[i];
                                SentItem.Move((Outlook.Folder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail));
                                StrSentItems += SentItem.Subject + " - " + SentItem.SentOn.ToString() + "\r\n";
                            }
                            else if (SentItems[i] is Outlook.MeetingItem)
                            {
                                Outlook.MeetingItem SentItem = (Outlook.MeetingItem)SentItems[i];
                                SentItem.Move((Outlook.Folder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail));
                                StrSentItems += SentItem.Subject + " - " + SentItem.SentOn.ToString() + "\r\n";
                            }
                            else
                                continue;
                        }
                    }
                    WriteLog(LOGTYPE.LT_INFORMATION, StrSentItems);
                }
            }
            catch (Exception ex)
            {
                WriteLog(LOGTYPE.LT_ERROR, "Error processing outbox: " + ex.Message);
            }
        }

        private void CheckAndMoveSentItemsThread()
        {
            while(config.BackgroundMonitor == true)
            {
                CheckAndMoveSentItems();
                Int32 Interval = 60;
                if(config.BackgroundInterval.HasValue == true)
                {
                    Interval = config.BackgroundInterval.Value;
                }
                Thread.Sleep(Interval * 1000);
            }
        }

        private Int32 GetHash(List<object> MailItems)
        {
            Int32 hash = 0;

            foreach(object MailItem in MailItems)
            {
                string EntryID;
                if (MailItem is Outlook.MailItem)
                    EntryID = ((Outlook.MailItem)MailItem).EntryID;
                else if (MailItem is Outlook.MeetingItem)
                    EntryID = ((Outlook.MeetingItem)MailItem).EntryID;
                else
                    continue;
                for(Int32 i = 0; i < (EntryID.Length / 2) / 4; i++)
                {
                    hash ^= Convert.ToInt32(EntryID.Substring((i * 2) * 4, 4), 16);
                    hash = (hash << 1) | ((hash >> 31) & 1);
                }
            }

            return hash;
        }

        private void WriteLog(LOGTYPE LogType, string LogString)
        {
            try
            {
                CheckLogPaths();
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + ": " + LogString);
                string LogPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Outbox Monitor\\Logs";
                
                if(config.LogLevel.HasValue == true && LogType > config.LogLevel)
                {
                    return;
                }
                
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

        private void LoadConfig()
        {
            try
            {
                CheckConfig();
                string LocalAppdataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                Config? loadedConfig = JsonSerializer.Deserialize<Config>(File.ReadAllText(LocalAppdataPath + "\\Outbox Monitor\\Config.json"));
                if (loadedConfig != null)
                {
                    config = loadedConfig;
                }
            }
            catch (Exception ex)
            {
                WriteLog(LOGTYPE.LT_ERROR, "Error loading configuration: " + ex.Message);
            }
        }

        private void CheckConfig()
        {
            string LocalAppdataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (Directory.Exists(LocalAppdataPath + "\\Outbox Monitor") == false)
            {
                Directory.CreateDirectory(LocalAppdataPath + "\\Outbox Monitor");
            }
            if (File.Exists(LocalAppdataPath + "\\Outbox Monitor\\Config.json") == false)
            {
                Config newConfig = new Config
                {
                    BackgroundMonitor = true,
                    BackgroundInterval = 60,
                    LogLevel = LOGTYPE.LT_INFORMATION,
                    LogOnly = false
                };
                string jsonString = JsonSerializer.Serialize(newConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(LocalAppdataPath + "\\Outbox Monitor\\Config.json", jsonString);
            }
            else
            {
                Config? loadedConfig = JsonSerializer.Deserialize<Config>(File.ReadAllText(LocalAppdataPath + "\\Outbox Monitor\\Config.json"));
                if (loadedConfig != null)
                {
                    bool changed = false;
                    if (loadedConfig.BackgroundMonitor.HasValue == false)
                    {
                        loadedConfig.BackgroundMonitor = true;
                        changed = true;
                    }
                    if ((loadedConfig.BackgroundInterval.HasValue == false) || loadedConfig.BackgroundInterval < 1)
                    {
                        loadedConfig.BackgroundInterval = 60;
                        changed = true;
                    }
                    if((loadedConfig.LogLevel.HasValue == false) || loadedConfig.LogLevel > LOGTYPE.LT_NONE || loadedConfig.LogLevel < 0)
                    {
                        loadedConfig.LogLevel = LOGTYPE.LT_INFORMATION;
                        changed = true;
                    }
                    if (loadedConfig.LogOnly.HasValue == false)
                    {
                        loadedConfig.LogOnly = false;
                        changed = true;
                    }
                    if (changed == true)
                    {
                        SaveConfig(loadedConfig);
                        WriteLog(LOGTYPE.LT_INFORMATION, "Saved configuration: " + JsonSerializer.Serialize(loadedConfig, new JsonSerializerOptions { WriteIndented = true }));
                    }
                }
            }
        }

        private void SaveConfig(Config SaveConfig = null)
        {
            try
            {
                if (SaveConfig != null)
                {
                    config = SaveConfig;
                }
                string LocalAppdataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string jsonString = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(LocalAppdataPath + "\\Outbox Monitor\\Config.json", jsonString);
            }
            catch (Exception ex)
            {
                WriteLog(LOGTYPE.LT_ERROR, "Error saving configuration: " + ex.Message);
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
