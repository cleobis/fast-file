﻿using System;
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

namespace QuickFile
{
    public partial class ThisAddIn
    {
        internal TaskPaneControlWrapper taskPaneControl;
        internal Microsoft.Office.Tools.CustomTaskPane customTaskPane;
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

                folderTree = new FolderWrapper(root, foldersCollection, null);
                foldersCollection = new ObservableCollection<FolderWrapper>(); //Change to incremental update later *****

            }
            else
            {
                // Clear collection
                while (foldersCollection.Count > 0)
                {
                    foldersCollection.RemoveAt(foldersCollection.Count - 1);
                }
                MessageBox.Show("Rebuilding");
            }

            // Build or re-build collection
            foreach (FolderWrapper fw in folderTree.Flattened())
            {
                foldersCollection.Add(fw);
            }

            if (!(taskPaneControl is null))
            {
                taskPaneControl.taskPaneControl.RefreshSelection();
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

    public class TaskPaneContext
    {
        private Outlook.Explorer explorer;
        private Outlook.Inspector inspector;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        private TaskPaneControl control;

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
            }
        }

        private object explorerOrInspector
        {
            get { return explorer == null ? (object)inspector : (object)explorer; }
        }

        public bool Visible
        {
            get { return taskPane.Visible; }
            set { 
                taskPane.Visible = value;
                if (value)
                {
                    control.textBox.Focus();
                }
            }
        }

        public void CloseCallback()
        {
            if (taskPane != null)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
            }
            taskPane = null;

            Globals.ThisAddIn.TaskPaneContexts.Remove(explorerOrInspector);
            if (explorer is null)
            {
                // Free inspector
                ((Outlook.InspectorEvents_Event)inspector).Close -= new Outlook.InspectorEvents_CloseEventHandler(CloseCallback);
                inspector = null;
            }
            else
            {
                // Free Explorer
                ((Outlook.ExplorerEvents_10_Event)explorer).Close -= new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_CloseEventHandler(CloseCallback);
            }
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
        }
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
        public String displayName;
        private Outlook.Folders folders; // Have to retain this reference or the events get garbage collected
        public FolderWrapper parent;
        public List<FolderWrapper> children;
        public ObservableCollection<FolderWrapper> collection;

        public FolderWrapper(Outlook.Folder folder, ObservableCollection<FolderWrapper> collection, FolderWrapper parent = null)
        {
            this.folder = folder;
            this.parent = parent;
            this.collection = collection;
            this.folders = folder.Folders;

            int depth = 0;
            var p = folder.Parent;
            while (p is Outlook.Folder)
            {
                depth += 1;
                p = (p as Outlook.Folder).Parent;
            }

            this.displayName = string.Concat(Enumerable.Repeat(" - ", depth)) + folder.Name;
            this.displayName = folder.FolderPath;

            //collection.Add(this);

            children = new List<FolderWrapper>(folders.Count);
            foreach (Outlook.Folder child in folders)
            {
                var fw = new FolderWrapper(child as Outlook.Folder, collection, this);
                children.Add(fw);
            }

            // Listeners
            folders.FolderAdd += new Outlook.FoldersEvents_FolderAddEventHandler(Folders_FolderAdd);
            folders.FolderChange += new Outlook.FoldersEvents_FolderChangeEventHandler(Folders_FolderChange);
            folders.FolderRemove += new Outlook.FoldersEvents_FolderRemoveEventHandler(Folders_FolderRemove);
        }

        public override String ToString()
        {
            return displayName + "(" + folder.Class + ")";
        }

        public void Folders_FolderAdd(Outlook.MAPIFolder new_folder)
        {
            FolderWrapper fw = new FolderWrapper(new_folder as Outlook.Folder, collection, this);
            //children.Insert(0,fw);
            //collection.Insert(collection.IndexOf(this) + 1, fw);

            MessageBox.Show(String.Format("Added {0} folder to {1}.", new_folder.Name, this.folder.Name));

            Globals.ThisAddIn.UpdateFolderList();
        }

        public void Folders_FolderChange(Outlook.MAPIFolder folder)
        {
            // Rename, Add child, or delete child.
            // ********** DEAL WITH RENAME ***************
            MessageBox.Show(String.Format(
                //"Changed {0} folder. ", folder.Name));
                "Changed {0} folder in {1}. ", folder.Name, this.folder.Name));
        }

        public void Folders_FolderRemove()
        {
            //MessageBox.Show("Removed a folder.");
            MessageBox.Show(String.Format("Removed a folder from {0}.", this.folder.Name));

            // Temp list of remaining folder for search convenience.
            var remainingFolderIds = new List<String>(folders.Count);
            foreach (Outlook.Folder f in folders)
            {
                remainingFolderIds.Add(f.EntryID);
            }

            for (int i = 0; i < children.Count; i++)
            {
                if (!remainingFolderIds.Contains(children[i].folder.EntryID))
                {
                    //RemoveChild(i);
                    children.RemoveAt(i);
                    return;
                }
            }
            MessageBox.Show("Unable to find deleted folder");

            Globals.ThisAddIn.UpdateFolderList();
        }

        /*public void RemoveChild(int i)
        {
            var child = children[i];
            for (int j = child.children.Count - 1; j >= 0; j--)
            {
                child.RemoveChild(j);
            }
            collection.Remove(child);
            children.RemoveAt(i);
        }*/

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

        private bool LeftCtrl = false;
        private bool RightCtrl = false;
        private bool LeftAlt = false;
        private bool RightAlt = false;
        private bool LeftShift = false;
        private bool RightShift = false;

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
                Debug.WriteLine("Already attached.");
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
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
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
            // Calls once n = 3 (call to peek), then n = 0 (calls to get the key)

            var key = KeyInterop.KeyFromVirtualKey((int)wParam);
            var flags = new KeystrokeFlags(lParam);
            switch (key)
            {
                case Key.LeftCtrl:
                    LeftCtrl = flags.IsDown;
                    break;
                case Key.RightCtrl:
                    RightCtrl = flags.IsDown;
                    break;
                case Key.LeftShift:
                    LeftShift = flags.IsDown;
                    break;
                case Key.RightShift:
                    RightShift = flags.IsDown;
                    break;
                case Key.LeftAlt:
                    LeftAlt = flags.IsDown;
                    break;
                case Key.RightAlt:
                    RightAlt = flags.IsDown;
                    break;
                case Key.V:
                    if ((LeftCtrl || RightCtrl) && (LeftShift || RightShift) && !(LeftAlt || RightAlt) && flags.IsDown)
                    {
                        ShowGUI();
                        return IntPtr.Zero + 1;
                    }
                    break;
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private void ShowGUI()
        {
            var window = Globals.ThisAddIn.Application.ActiveWindow();
            var context = Globals.ThisAddIn.TaskPaneContexts[window];
            if (context != null)
            {
                context.Visible = true;
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
}