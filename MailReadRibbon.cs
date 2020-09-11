using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;

namespace QuickFile
{
    public partial class MailReadRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Inspector inspector = (Outlook.Inspector)e.Control.Context;
            TaskPaneContext taskPaneContext = Globals.ThisAddIn.TaskPaneContexts[inspector];
            taskPaneContext.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void guessButton_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Inspector inspector = (Outlook.Inspector)e.Control.Context;
            TaskPaneContext taskPaneContext = Globals.ThisAddIn.TaskPaneContexts[inspector];
            taskPaneContext.MoveSelectedItemToBest();
        }
    }
}
