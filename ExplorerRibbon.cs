using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using System.Diagnostics;
using System.Windows;

namespace QuickFile
{
    public partial class ExplorerRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Explorer explorer = (Outlook.Explorer)e.Control.Context;
                TaskPaneContext taskPaneContext = Globals.ThisAddIn.TaskPaneContexts[explorer];
                taskPaneContext.MoveSelectedItemToBest();
            }
            catch (Exception err)
            {
                MessageBox.Show("Unexpected error processing button.\n" + err.Message, "Fast File Error");
                Debug.WriteLine("Unexpected error processing button.\n" + err.Message + $"\n{err}");
            }
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Explorer explorer = (Outlook.Explorer)e.Control.Context;
                TaskPaneContext taskPaneContext = Globals.ThisAddIn.TaskPaneContexts[explorer];
                taskPaneContext.Visible = ((RibbonToggleButton)sender).Checked;
            }
            catch (Exception err)
            {
                MessageBox.Show("Unexpected error processing button.\n" + err.Message, "Fast File Error");
                Debug.WriteLine("Unexpected error processing button.\n" + err.Message + $"\n{err}");
            }
        }
    }
}
