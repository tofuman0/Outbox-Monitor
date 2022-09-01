using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace Outbox_Monitor
{
    public partial class OutboxMonitorRibbon
    {
        private void OutboxMonitorRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnProcessOutboxItems_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckAndMoveSentItems();
        }
    }
}
