using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Security.AccessControl;
using System.Collections.ObjectModel;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.Threading;
using System.Windows.Threading;

namespace QuickFile
{
    public partial class ThisAddIn
    {
        internal ObservableCollection<FolderWrapper> foldersCollection;
        internal FolderWrapper folderTree;
        
        public Dictionary<object, TaskPaneContext> TaskPaneContexts = new Dictionary<object, TaskPaneContext>();
        private Outlook.Inspectors inspectors; // So event isn't garbage collected.
        private Outlook.Explorers explorers; // So event isn't garbage collected.

        private InterceptKeys interceptKeys;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            UpdateFolderList();

            interceptKeys = new InterceptKeys();
            interceptKeys.Attach();

            // Attach task panes to inspectors
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(NewInspector);
            foreach (Outlook.Inspector inspector in inspectors)
            {
                NewInspector(inspector);
            }

            // Attach task panes to explorers
            explorers = this.Application.Explorers;
            explorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(NewExplorer);
            foreach (Outlook.Explorer inspector in explorers)
            {
                NewExplorer(inspector);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void NewInspector(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.MailItem)
            {
                TaskPaneContexts.Add(inspector, new TaskPaneContext(inspector));
            }
        }

        private void NewExplorer(Outlook.Explorer explorer)
        {
            TaskPaneContexts.Add(explorer, new TaskPaneContext(explorer));
        }

        public void UpdateFolderList()
        {

            if (folderTree is null)
            {
                Outlook.Folder root = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                // or loop over Application.Session.Stores

                folderTree = new FolderWrapper(root, null, GetDefaultFolders(false));
                foldersCollection = new ObservableCollection<FolderWrapper>(); //Change to incremental update later

            }
            else
            {
                // Clear collection
                while (foldersCollection.Count > 0)
                {
                    foldersCollection.RemoveAt(foldersCollection.Count - 1);
                }
            }

            // Build or re-build collection
            foreach (FolderWrapper fw in folderTree.Flattened())
            {
                foldersCollection.Add(fw);
            }

            // Update selection on control
            foreach (var pair in TaskPaneContexts)
            {
                pair.Value.Refresh();
            }
        }

        public List<Outlook.Folder> GetDefaultFolders(bool includeInbox)
        {
            var folders = new List<Outlook.Folder>();
            
            // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.oldefaultfolders?view=outlook-pia
            // https://stackoverflow.com/questions/972307/how-to-loop-through-all-enum-values-in-c
            foreach (var folderType in EnumUtil.GetValues<Outlook.OlDefaultFolders>())
            {
                switch (folderType)
                {
                    // Folders to suppress
                    case Outlook.OlDefaultFolders.olFolderCalendar:
                    case Outlook.OlDefaultFolders.olFolderConflicts:
                    case Outlook.OlDefaultFolders.olFolderContacts:
                    case Outlook.OlDefaultFolders.olFolderDrafts:
                    case Outlook.OlDefaultFolders.olFolderJournal:
                    case Outlook.OlDefaultFolders.olFolderLocalFailures:
                    case Outlook.OlDefaultFolders.olFolderNotes:
                    case Outlook.OlDefaultFolders.olFolderOutbox:
                    case Outlook.OlDefaultFolders.olFolderRssFeeds:
                    case Outlook.OlDefaultFolders.olFolderSentMail:
                    case Outlook.OlDefaultFolders.olFolderServerFailures:
                    case Outlook.OlDefaultFolders.olFolderSyncIssues:
                    case Outlook.OlDefaultFolders.olFolderTasks:
                    case Outlook.OlDefaultFolders.olFolderToDo:
                        try
                        {
                            folders.Add(Application.Session.DefaultStore.GetDefaultFolder(folderType) as Outlook.Folder);
                        } catch (COMException err) {
                            if (err.ErrorCode != -2147221233 // folder not found
                                && err.ErrorCode != 0x8004060E) // Exchange connection required.
                            {
                                throw err;
                            }
                        }
                        break;

                    // Folders to allow
                    case Outlook.OlDefaultFolders.olFolderDeletedItems:
                    case Outlook.OlDefaultFolders.olFolderJunk:
                    case Outlook.OlDefaultFolders.olPublicFoldersAllPublicFolders:
                        break;

                    // Depends on input argument
                    case Outlook.OlDefaultFolders.olFolderInbox:
                        if (includeInbox)
                        {
                            try
                            {
                                folders.Add(Application.Session.DefaultStore.GetDefaultFolder(folderType) as Outlook.Folder);
                            }
                            catch (COMException err)
                            {
                                if (err.ErrorCode != 2147221233) // folder not found
                                {
                                    throw err;
                                }
                            }
                        }
                        break;

                    // Folders to do nothing
                    case Outlook.OlDefaultFolders.olFolderManagedEmail:
                    case Outlook.OlDefaultFolders.olFolderSuggestedContacts:

                        break;
                }
            }
            return folders;
        }

        public void RehookKeyboard()
        {
            interceptKeys.Detach();
            interceptKeys.Attach();
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

    public class TaskPaneContext
    {
        public readonly Outlook.Explorer explorer;
        public readonly Outlook.Inspector inspector;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        private TaskPaneControl control;
        private FolderWrapper _bestFolderWrapper;
        private bool guessBestFolderQueued = false;

        public TaskPaneContext(Outlook.Explorer explorer) : this(explorer, null) { }
        public TaskPaneContext(Outlook.Inspector inspector) : this(null, inspector) { }
        private TaskPaneContext(Outlook.Explorer explorer, Outlook.Inspector inspector)
        {
            this.explorer = explorer;
            this.inspector = inspector;

            var wrapper = new TaskPaneControlWrapper();
            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(wrapper, "Quick Move", explorerOrInspector);
            taskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);
            control = wrapper.taskPaneControl;
            control.SetUp();
            control.taskPaneContext = this;

            if (this.explorer is null)
            {
                // Init inspector
                ((Outlook.InspectorEvents_Event)this.inspector).Close += new Outlook.InspectorEvents_CloseEventHandler(CloseCallback);
            }
            else
            {
                // Init explorer
                ((Outlook.ExplorerEvents_10_Event)this.explorer).Close += new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_CloseEventHandler(CloseCallback);
                ((Outlook.ExplorerEvents_10_Event)this.explorer).SelectionChange += new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
            }

            GuessBestFolderAsync();
        }

        internal object explorerOrInspector
        {
            get { return explorer == null ? (object)inspector : (object)explorer; }
        }
        
        public void UpdateBestFolderWrapper()
        {
            // Called from inspector ribbon load event.
            UpdateBestFolderWrapper(_bestFolderWrapper);
        }
        private void UpdateBestFolderWrapper(FolderWrapper value)
        {
            _bestFolderWrapper = value;
            // Update Panel
            control.RefreshSelection(_bestFolderWrapper); // null will be ignored.

            // Update Ribbon
            RibbonButton button = null;
            if (explorer != null)
            {
                button = Globals.Ribbons[explorer].ExplorerRibbon?.guessButton;
            }
            else
            {
                // The ribbon is not loaded until after the New Inspector event.
                button = Globals.Ribbons[inspector].MailReadRibbon?.guessButton;
            }
            if (button != null)
            {
                if (_bestFolderWrapper?.folder == null)
                {
                    button.Label = "Move";
                    button.Enabled = false;
                }
                else
                {
                    button.Label = _bestFolderWrapper.folder.Name;
                    button.Enabled = true;
                }
            }
        }
        
        public bool Visible
        {
            get { return taskPane.Visible; }
            set {
                bool changed = taskPane.Visible != value;
                taskPane.Visible = value;
                if (value)
                {
                    control.textBox.Focus();

                    // Fix escape key capturing
                    var id = GetForegroundWindow();
                    SetForegroundWindow(GetDesktopWindow());
                    SetForegroundWindow(id);
                }
                else
                {
                    if (changed)
                    {
                        // Fix focus not returned to mesage list in explorer
                        var id = GetForegroundWindow();
                        SetForegroundWindow(GetDesktopWindow());
                        SetForegroundWindow(id);
                    }
                }
            }
        }

        public void MoveSelectedItem(Outlook.Folder folder)
        {
            foreach (Outlook.MailItem mailItem in GetSelectedMailItems())
            {
                mailItem.Move(folder);
            }
        }
        
        public void MoveSelectedItemToBest()
        {
            if (_bestFolderWrapper is null)
            {
                Debug.WriteLine("No best folder.");
            }
            else
            {
                MoveSelectedItem(_bestFolderWrapper.folder);
            }
        }

        public void Refresh()
        {
            control.RefreshSelection();
        }

        internal void GuessBestFolderAsync()
        {
            UpdateBestFolderWrapper(null);
            // The Explorer selection changed event fires twice if the reading 
            // pane is open. By pushing our response back into the message queue,
            // we can consolidate and only update once. Using Background priority
            // means that the reading pane render will happen first.
            if (!guessBestFolderQueued)
            {
                guessBestFolderQueued = true;
                Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() =>
                {
                    guessBestFolderQueued = false;
                    try
                    {
                        GuessBestFolder();
                    }
                    catch (Exception err)
                    {
                        MessageBox.Show("Unexpected error processing GuessBestFolderAsync.\n" + err.Message, "Fast File Error");
                        Debug.WriteLine("Unexpected error processing GuessBestFolderAsync.\n" + err.Message + $"\n{err}");
                    }
                }), DispatcherPriority.Background);
            }
        }
        
        internal void GuessBestFolder()
        {
            // Which folder contains the most messages from the conversation?
            Dictionary<String, Tuple<Outlook.Folder, int>> folderVotes = new Dictionary<String, Tuple<Outlook.Folder, int>>();
            void processItem(Outlook.MailItem mailItem)
            {
                Outlook.Conversation conv = null;
                try
                {
                    conv = mailItem.GetConversation();
                }
                catch (COMException err)
                {
                    //It's supposed to return null if there is no converstaion but actually throws and exception
                }

                if (conv != null)
                {
                    // Obtain root items and enumerate the conversation. 
                    Outlook.SimpleItems simpleItems = conv.GetRootItems();
                    EnumerateConversation(simpleItems, conv);
                }
            }
            void EnumerateConversation(Outlook.SimpleItems items, Outlook.Conversation conversation)
            {
                if (items.Count > 0)
                {
                    foreach (object myItem in items)
                    {
                        if (myItem is Outlook.MailItem)
                        {
                            Outlook.MailItem mailItem = myItem as Outlook.MailItem;
                            Outlook.Folder inFolder = mailItem.Parent as Outlook.Folder;

                            if (!folderVotes.TryGetValue(inFolder.EntryID, out Tuple<Outlook.Folder, int> value))
                            {
                                value = new Tuple<Outlook.Folder, int>(inFolder, 0);
                            }
                            folderVotes[inFolder.EntryID] = new Tuple<Outlook.Folder, int>(inFolder, value.Item2 + 1);
                        }
                        // Continue recursion. 
                        Outlook.SimpleItems children;
                        try
                        {
                            children = conversation.GetChildren(myItem);
                        }
                        catch (COMException err)
                        {
                            var subject = myItem is Outlook.MailItem ? (myItem as Outlook.MailItem).Subject : "<Unknown item>";
                            Debug.WriteLine("Unable to get conversation children for " + subject);
                            // I see this with Drafts, meeting invites, and other times.
                            continue;
                        }
                        EnumerateConversation(children, conversation);
                    }
                }
            }
            foreach (Outlook.MailItem mailItem in GetSelectedMailItems()) {
                processItem(mailItem);
            }

            // Remove distracting folders from consideration.
            var folderBlacklist = Globals.ThisAddIn.GetDefaultFolders(false);
            if (explorer != null)
            {
                folderBlacklist.Add(explorer.CurrentFolder as Outlook.Folder);
            }
            else // inspector 
            {
                if (inspector.CurrentItem is Outlook.MailItem)
                {
                    var mailItem = inspector.CurrentItem as Outlook.MailItem;
                    if (mailItem.Parent is Outlook.Folder)
                    {
                        folderBlacklist.Add(mailItem.Parent as Outlook.Folder);
                    }
                }
            }

            // Select best folder
            var sortedFolders = folderVotes.OrderBy(key => -key.Value.Item2);
            Outlook.Folder bestFolder = null;
            foreach (var v in sortedFolders)
            {
                Outlook.Folder folder = v.Value.Item1;
                if (folderBlacklist.FindIndex(f => f.EntryID == folder.EntryID) >= 0)
                {
                    // on blacklist
                    continue;
                }
                bestFolder = folder;
                break;
            }

            // Return folder wrapper
            FolderWrapper best = null;
            if (bestFolder != null)
            {
                try
                {
                    best = Globals.ThisAddIn.foldersCollection.Single(fw => fw.folder.EntryID == bestFolder.EntryID);
                }
                catch (InvalidOperationException)
                {
                    Debug.WriteLine("Unable to find folder " + bestFolder.Name + ".");
                }
            }

            UpdateBestFolderWrapper(best);
        }

        public IEnumerable<Outlook.MailItem> GetSelectedMailItems()
        {
            if (inspector != null)
            {
                if (inspector.CurrentItem is Outlook.MailItem)
                {
                    yield return inspector.CurrentItem as Outlook.MailItem;
                }
            }
            else // (explorer != null)
            {
                Outlook.Selection headers = null;
                try
                {
                    headers = explorer.Selection.GetSelection(Outlook.OlSelectionContents.olConversationHeaders);
                }
                catch (COMException err)
                {
                    // ^ failed once when moving only part of the message.
                    Debug.WriteLine("Error with GetSelection().");
                }


                if (headers != null && headers.Count > 0)
                {
                    // If they are in conversation view, need to iterate through the conversations in case they have the header selected. Only returns items in the current folder which is what we want.
                    foreach (Outlook.ConversationHeader header in headers)
                    {
                        // System.Runtime.InteropServices.COMException
                        // Message = The operation failed.
                        // after ?moving? a conversation and definitely after ?deleting? a conversation.
                        Outlook.SimpleItems items = null;
                        try
                        {
                            items = header.GetItems();
                        }
                        catch (COMException err)
                        {
                            // Seen after move sometimes
                            Debug.WriteLine("COMException in header.GetItems(). " + err.Message);
                            continue;
                        }
                        for (int i = 1; i <= items.Count; i++)
                        {
                            // Enumerate only MailItems in this example.
                            if (items[i] is Outlook.MailItem)
                            {
                                yield return items[i] as Outlook.MailItem;
                            }
                        }
                    }
                }
                else
                {
                    // If we are not in conversation view, process selection directly
                    Outlook.Selection selection = null;
                    try
                    {
                        selection = explorer.Selection;
                    }
                    catch (COMException err)
                    {
                        Debug.WriteLine("ComException in explorer.Selection." + err.Message);
                    }
                    if (selection != null)
                    {
                        for (int i = 1; i <= selection.Count; i++)
                        {
                            var selectionItem = selection[i];
                            if (selectionItem is Outlook.MailItem)
                            {
                                yield return selectionItem as Outlook.MailItem;
                            }
                        }
                    }
                }
            }
        }

        public void Explorer_SelectionChange()
        {
            try
            {
                GuessBestFolderAsync();
            }
            catch (Exception err)
            {
                MessageBox.Show("Unexpected error processing Selection Change.\n" + err.Message,"Fast File Error");
                Debug.WriteLine("Unexpected error processing Selection Change.\n" + err.Message + $"\n{err}");
            }
        }


        public void CloseCallback()
        {
            
            Globals.ThisAddIn.TaskPaneContexts.Remove(explorerOrInspector);
            if (explorer is null)
            {
                // Free inspector
                ((Outlook.InspectorEvents_Event)inspector).Close -= new Outlook.InspectorEvents_CloseEventHandler(CloseCallback);
            }
            else
            {
                // Free Explorer
                ((Outlook.ExplorerEvents_10_Event)explorer).Close -= new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_CloseEventHandler(CloseCallback);
                ((Outlook.ExplorerEvents_10_Event)this.explorer).SelectionChange -= new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
            }
            
            if (taskPane != null)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
            }
            taskPane = null;
        }

        void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            RibbonToggleButton toggleButton = null;
            if (explorer != null)
            {
                toggleButton = Globals.Ribbons[explorer].ExplorerRibbon.toggleButton1;
            }
            else
            {
                toggleButton = Globals.Ribbons[inspector].MailReadRibbon.toggleButton1;
            }
            toggleButton.Checked = taskPane.Visible;
            if (!taskPane.Visible)
            {
                control.textBox.Text = "";
            }
        }

        [DllImport("user32.dll", SetLastError = false)]
        static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
    }

    class FolderWrapper
    {
        /* Object for keeping track of the folder hierarchy

           Folder Change Events
           ====================
           Given a collection of folders in Outlook, several events are raised when folders in that collection change:
           Folders.FolderAdd is raised on a Folders collection when a new folder is added. Outlook passes a folder parameter of type MAPIFolder representing the newly added folder.
           Folders.FolderRemove is raised on a Folders collection when a folder is deleted.
           Folders.FolderChange is raised on a Folders collection when a folder is changed. Examples of changes include when the folder is renamed or when the number of items in the folder changes. Outlook passes a folder parameter of type MAPIFolder representing the folder that has changed.

           When a subfolder is deleted (moved to Trash), the following events are called in this order:
            - Change called Trash
            - Added called on Trash
            - Removed called on old parent
            - Changed called on old parent
        */
        public Outlook.Folder folder;
        private Outlook.Folders folders; // Have to retain this reference or the events get garbage collected
        public FolderWrapper parent;
        public List<FolderWrapper> children;
        private int depth;
        public bool stale = false;

        public FolderWrapper(Outlook.Folder folder, FolderWrapper parent = null, List<Outlook.Folder> blacklist = null)
        {
            this.folder = folder;
            this.parent = parent;
            this.folders = folder.Folders;

            depth = 0;
            var p = folder.Parent;
            while (p is Outlook.Folder)
            {
                depth += 1;
                p = (p as Outlook.Folder).Parent;
            }

            children = new List<FolderWrapper>(folders.Count);
            foreach (Outlook.Folder child in folders)
            {
                if (blacklist != null && blacklist.FindIndex(f => f.EntryID == child.EntryID) >= 0)
                {
                    continue;
                }
                var fw = new FolderWrapper(child, this, blacklist);
                children.Add(fw);
            }

            // Listeners
            folders.FolderAdd += new Outlook.FoldersEvents_FolderAddEventHandler(Folders_FolderAdd);
            folders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(Folders_FolderChange);
            folders.FolderRemove += new Outlook.FoldersEvents_FolderRemoveEventHandler(Folders_FolderRemove);
        }

        public override String ToString()
        {
            return DisplayName;
        }

        public String DisplayName
        {
            get { return string.Concat(Enumerable.Repeat("  ", depth)) + folder.Name; }
        }

        public void Folders_FolderAdd(Outlook.MAPIFolder new_folder)
        {
            try
            {
                Debug.WriteLine("FolderAdd Starting");

                FolderWrapper fw = new FolderWrapper(new_folder as Outlook.Folder, this, Globals.ThisAddIn.GetDefaultFolders(false));
                //children.Insert(0,fw);
                //collection.Insert(collection.IndexOf(this) + 1, fw);

                Globals.ThisAddIn.UpdateFolderList();

                Debug.WriteLine("FolderAdd Done.");
            }
            catch (Exception err)
            {
                MessageBox.Show("Unexpected error processing FolderAdd.\n" + err.Message,"Fast File Error");
                Debug.WriteLine("Unexpected error processing FolderAdd.\n" + err.Message + $"\n{err}");
            }
}

        public void Folders_FolderChange(Outlook.MAPIFolder folder)
        {
            try
            { 

            }
            catch (Exception err)
            {
                MessageBox.Show("Unexpected error processing FolderChange.\n" + err.Message,"Fast File Error");
                Debug.WriteLine("Unexpected error processing FolderChange.\n" + err.Message + $"\n{err}");
            }
    // Rename, Add child, or delete child.
    // TODO: THIS IS TOO SLOW TO HAVE ENABLED. NEED TO FIX.
    //Globals.ThisAddIn.UpdateFolderList();
}

        public void Folders_FolderRemove()
        {
            try
            {
                Debug.WriteLine("FolderRemove Starting");

                // Temp list of remaining folder for search convenience.
                var remainingFolderIds = new List<String>(folders.Count);
                foreach (Outlook.Folder f in folders)
                {
                    remainingFolderIds.Add(f.EntryID);
                }

                for (int i = children.Count; i >= 0; i--)
                {
                    if (!remainingFolderIds.Contains(children[i].folder.EntryID))
                    {
                        children.RemoveAt(i);
                        return;
                    }
                }

                Globals.ThisAddIn.UpdateFolderList();

                Debug.WriteLine("FolderRemove Done");
            }
            catch (Exception err)
            {
                MessageBox.Show("Unexpected error processing FolderRemove.\n" + err.Message, "Fast File Error");
                Debug.WriteLine("Unexpected error processing FolderRemove.\n" + err.Message + $"\n{err}");
            }
        }

        public IEnumerable<FolderWrapper> Flattened()
        {
            yield return this;
            foreach (var child in children)
            {
                foreach (var i in child.Flattened())
                {
                    yield return i;
                }
            }
        }
    }

    class InterceptKeys
    {
        // http://web.archive.org/web/20190828074433/https://blogs.msdn.microsoft.com/toub/2006/05/03/low-level-keyboard-hook-in-c/

        private const int WH_KEYBOARD_LL = 13;
        private const int WH_KEYBOARD = 2;
        private LowLevelKeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;

        public InterceptKeys()
        {
            _proc = HookCallback;
        }
        ~InterceptKeys()
        {
            Detach();
        }

        public bool Attach()
        {
            if (_hookID != IntPtr.Zero)
            {
                return false;
            }
            _hookID = SetHook(_proc);
            return _hookID != IntPtr.Zero;
            ;
        }

        public void Detach()
        {
            if (_hookID != IntPtr.Zero)
            {
                bool check = UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
                Debug.WriteLine("Detaching hook. Return was " + check);
            }
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            using (ProcessThread thread = curProcess.Threads[0])
            {
                return SetWindowsHookEx(WH_KEYBOARD, proc, IntPtr.Zero, (uint)curProcess.Threads[0].Id);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            // Calls twice n = 3 (call to peek), then once n = 0 (calls to get the key). Calls for key down, key repeat, and key up
            //
            // Seeing issues where the hook stops responding after a while. Suspect the OS is unhooking due to slow execution. Trying to mitigate by pushing the actual work back on the Dispatcher queue rather than doing within the HookCallback.
            if (nCode < 0)
            {
                return CallNextHookEx(_hookID, nCode, wParam, lParam);
            }

            var key = KeyInterop.KeyFromVirtualKey((int)wParam);
            var flags = new KeystrokeFlags(lParam);

            // We only care about the first key down. We don't want repeats or key ups.
            // nCode == 0 is the actual key press, not the peak which could happen multiple times.
            if (!(flags.WasUp && flags.IsDown && nCode == 0))
            {
                return CallNextHookEx(_hookID, nCode, wParam, lParam);
            }

            switch (key)
            {
                case Key.D1:
                    Debug.WriteLine("Got D1. Modifiers = " + Keyboard.Modifiers + ", key = " + key + ", flags = " + flags);
                    if (Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
                    {
                        // Ctrl+Shfit+1 => Show GUI
                        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => {
                            try
                            {
                                var context = GetActiveContext();
                                if (context != null)
                                {
                                    context.Visible = true;
                                }
                            }
                            catch (Exception err)
                            {
                                MessageBox.Show("Unexpected error processing Ctrl+Shift+1.\n" + err.Message, "Fast File Error");
                                Debug.WriteLine("Unexpected error processing Ctrl+Shift+1.\n" + err.Message + $"\n{err}");
                            }
                        }));
                        return IntPtr.Zero + 1;
                    }
                    else if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        // Ctrl+1 => Move selected item to best guess
                        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => {
                            try
                            {
                                GetActiveContext()?.MoveSelectedItemToBest();
                            }
                            catch (Exception err)
                            {
                                MessageBox.Show("Unexpected error processing Ctrl+1.\n" + err.Message, "Fast File Error");
                                Debug.WriteLine("Unexpected error processing Ctrl+1.\n" + err.Message + $"\n{err}");
                            }
                    }));
                        return IntPtr.Zero + 1;
                    }
                    break;
                case Key.V:

                    Debug.WriteLine("Got v. Modifiers = " + Keyboard.Modifiers + ", key = " + key + ", flags = " + flags);
                    if (Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
                    {
                        // Ctrl+Shfit+V => Show GUI
                        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => {
                            try
                            {
                                var context = GetActiveContext();
                                if (context != null)
                                {
                                    context.Visible = true;
                                }
                            }
                            catch (Exception err)
                            {
                                MessageBox.Show("Unexpected error processing Ctrl+Shift+V.\n" + err.Message, "Fast File Error");
                                Debug.WriteLine("Unexpected error processing Ctrl+Shift+V.\n" + err.Message + $"\n{err}");
                            }
                        }));
                        return IntPtr.Zero + 1;
                    }
                    break;
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private TaskPaneContext GetActiveContext()
        {
            var window = Globals.ThisAddIn.Application.ActiveWindow();
            if (Globals.ThisAddIn.TaskPaneContexts.ContainsKey(window))
            {
                return Globals.ThisAddIn.TaskPaneContexts[window];
            }
            else
            {
                return null;
            }
        }

        internal struct KeystrokeFlags
        {
            /* 0-15
                The repeat count. The value is the number of times the keystroke is repeated as a result of the user's holding down the key.
                16-23
                The scan code. The value depends on the OEM.
                24
                Indicates whether the key is an extended key, such as a function key or a key on the numeric keypad. The value is 1 if the key is an extended key; otherwise, it is 0.
                25-28
                Reserved.
                29
                The context code. The value is 1 if the ALT key is down; otherwise, it is 0.
                30
                The previous key state. The value is 1 if the key is down before the message is sent; it is 0 if the key is up.
                31
                The transition state. The value is 0 if the key is being pressed and 1 if it is being released.
            */
            private long raw;
            public KeystrokeFlags(IntPtr _in)
            {
                raw = _in.ToInt64();
            }
            public int Repeat
            {
                get { return (int)(raw & 0x0000FFFF); }
                //set { raw = (uint)(raw & ~mask0 | (value << loc0) & mask0); }
            }
            public int ScanCode
            {
                get { return (int)(raw & 0x00FF0000) >> 16; }
            }
            public bool Alt
            {
                get { return Convert.ToBoolean(raw & 0x20000000); }
            }
            public bool WasDown
            {
                get { return Convert.ToBoolean(raw & 0x40000000); }
            }
            public bool WasUp
            {
                get { return !WasDown; }
            }
            public bool IsUp
            {
                get { return Convert.ToBoolean(raw & 0x80000000); }
            }
            public bool IsDown
            {
                get { return !IsUp; }
            }
            public override String ToString()
            {
                return String.Format("KeystrokeFlags: Repeat {0}, ScanCode {1}, Alt {2}, WasDown {3}, IsDown {4}.", Repeat, ScanCode, Alt, WasDown, IsDown);
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

    }

    public static class EnumUtil
    {
        public static IEnumerable<T> GetValues<T>()
        {
            return Enum.GetValues(typeof(T)).Cast<T>();
        }
    }
}
