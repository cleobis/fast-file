using ControlzEx.Standard;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Xaml.Behaviors.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:
// * Keyboard shortcuts
// * Subscribe to folder events
// * Multiple panels - https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-displaying-custom-task-panes-with-e-mail-messages-in-outlook?view=vs-2019#prerequisites
// * Guess item from conversation
// * Toolbar dropdown

namespace QuickFile
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class TaskPaneControl : UserControl
    {
        private ObservableCollection<FolderWrapper> foldersCollection;
        private FolderWrapper folderTree;

        public TaskPaneControl()
        {
            InitializeComponent();

            UpdateFolderList();

        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox senderTb = sender as TextBox;

            if (!(listBox is null))
            {
                CollectionViewSource.GetDefaultView(listBox.ItemsSource).Refresh();
                if (listBox.SelectedIndex == -1 && listBox.Items.Count > 0)
                {
                    listBox.SelectedIndex = 0;
                }
            }
        }

        void Explorer_SelectionChange()
        {
            var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            String str = "";

            str = String.Format("{0} items selected.\n\n", explorer.Selection.Count);
            for (int i = 1; i <= explorer.Selection.Count; i++)
            {
                var selection = explorer.Selection[i];
                if (selection is Outlook.MailItem)
                {
                    var mailItem = selection as Outlook.MailItem;
                    str += mailItem.Subject + "\n\n";

                    var conv = mailItem.GetConversation();
                    if (conv != null)
                    {
                        // Obtain root items and enumerate the conversation. 
                        Outlook.SimpleItems simpleItems = conv.GetRootItems();
                        foreach (object item in simpleItems)
                        {
                            if (item is Outlook.MailItem)
                            {
                                Outlook.MailItem mail = item as Outlook.MailItem;
                                Outlook.Folder inFolder = mail.Parent as Outlook.Folder;
                                string msg = mail.Subject + " in folder " + inFolder.Name;
                                str += msg + "\n";
                            }
                            // Call EnumerateConversation 
                            // to access child nodes of root items. 
                            str += EnumerateConversation(item, conv);
                        }
                    }
                }
                else
                {
                    str += "Not Mail item.\n\n";
                }
            }
            this.textBlock.Text = str;
        }

        String EnumerateConversation(object item, Outlook.Conversation conversation)
        {
            String str = "";
            Outlook.SimpleItems items = conversation.GetChildren(item);
            if (items.Count > 0)
            {
                foreach (object myItem in items)
                {
                    if (myItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = myItem as Outlook.MailItem;
                        Outlook.Folder inFolder = mailItem.Parent as Outlook.Folder;
                        string msg = mailItem.Subject + " in folder " + inFolder.Name;
                        str += msg + "\n";
                    }
                    // Continue recursion. 
                    str += EnumerateConversation(myItem, conversation);
                }
            }
            return str;
        }
        private bool FilterHelper(object obj)
        {
            var query = textBox.Text;
            if (string.IsNullOrEmpty(query))
            {
                return true;
            }
            else
            {
                return (obj.ToString().IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        internal class SortHelper : System.Collections.IComparer
        {
            public SortHelper(String _) { }
            public int Compare(object a, object b)
            {
                var fwa = a as FolderWrapper;
                var fwb = b as FolderWrapper;
                return String.Compare(fwa.folder.FolderPath, fwb.folder.FolderPath, true);

                // https://github.com/quicksilver/Quicksilver/blob/8be7395b795179cf51cf30ebf82779e0f9ba2138/Quicksilver/Code-QuickStepFoundation/QSSense.m
                /* 
                    #define MIN_ABBR_OPTIMIZE 0
                    #define IGNORED_SCORE 0.9
                    #define SKIPPED_SCORE 0.15



                    CGFloat QSScoreForAbbreviationWithRanges(CFStringRef str, CFStringRef abbr, id mask, CFRange strRange, CFRange abbrRange);

                    CGFloat QSScoreForAbbreviation(CFStringRef str, CFStringRef abbr, id mask) {
                    return QSScoreForAbbreviationWithRanges(str, abbr, mask, CFRangeMake(0, CFStringGetLength(str) ), CFRangeMake(0, CFStringGetLength(abbr)));
                    }

                    CGFloat QSScoreForAbbreviationWithRanges(CFStringRef str, CFStringRef abbr, id mask, CFRange strRange, CFRange abbrRange) {

                    if (!abbrRange.length)
                    return IGNORED_SCORE; //deduct some points for all remaining letters

                    if (abbrRange.length > strRange.length)
                    return 0.0;

                    // Create an inline buffer version of str.  Will be used in loop below
                    // for faster lookups.
                    CFStringInlineBuffer inlineBuffer;
                    CFStringInitInlineBuffer(str, &inlineBuffer, strRange);
                    CFLocaleRef userLoc = CFLocaleCopyCurrent();

                    CGFloat score = 0.0, remainingScore = 0.0;
                    NSInteger i, j;
                    CFRange matchedRange, remainingStrRange, adjustedStrRange = strRange;

                    for (i = abbrRange.length; i > 0; i--) { //Search for steadily smaller portions of the abbreviation
                    CFStringRef curAbbr = CFStringCreateWithSubstring (NULL, abbr, CFRangeMake(abbrRange.location, i) );
                    //terminality
                    //axeen
                    //        CFLocaleRef userLoc = CFLocaleCopyCurrent();
                    BOOL found = CFStringFindWithOptionsAndLocale(str, curAbbr,
                                                              CFRangeMake(adjustedStrRange.location, adjustedStrRange.length - abbrRange.length + i),
                                                              kCFCompareCaseInsensitive | kCFCompareDiacriticInsensitive | kCFCompareLocalized,
                                                              userLoc, &matchedRange);
                    CFRelease(curAbbr);
                    //        CFRelease(userLoc);

                    if (!found) {
                    continue;
                    }

                    if (mask) {
                    [mask addIndexesInRange:NSMakeRange(matchedRange.location, matchedRange.length)];
                    }


                    remainingStrRange.location = matchedRange.location + matchedRange.length;
                    remainingStrRange.length = strRange.location + strRange.length - remainingStrRange.location;

                    // Search what is left of the string with the rest of the abbreviation
                    remainingScore = QSScoreForAbbreviationWithRanges(str, abbr, mask, remainingStrRange, CFRangeMake(abbrRange.location + i, abbrRange.length - i) );

                    if (remainingScore) {
                    score = remainingStrRange.location-strRange.location;
                    // ignore skipped characters if is first letter of a word
                    if (matchedRange.location>strRange.location) {//if some letters were skipped
                        static CFCharacterSetRef wordSeparator = NULL;
                        if (!wordSeparator)
                        {
                          wordSeparator = CFCharacterSetCreateMutableCopy(NULL, CFCharacterSetGetPredefined(kCFCharacterSetWhitespace));
                          CFCharacterSetAddCharactersInString((CFMutableCharacterSetRef)wordSeparator, (CFStringRef)@".");
                        }
                        static CFCharacterSetRef uppercase = NULL;
                        if (!uppercase) uppercase = CFCharacterSetGetPredefined(kCFCharacterSetUppercaseLetter);
                        if (CFCharacterSetIsCharacterMember(wordSeparator, CFStringGetCharacterFromInlineBuffer(&inlineBuffer, matchedRange.location-1) )) {
                            for (j = matchedRange.location-2; j >= (NSInteger) strRange.location; j--) {
                                if (CFCharacterSetIsCharacterMember(wordSeparator, CFStringGetCharacterFromInlineBuffer(&inlineBuffer, j) )) score--;
                                else score -= SKIPPED_SCORE;
                            }
                        } else if (CFCharacterSetIsCharacterMember(uppercase, CFStringGetCharacterFromInlineBuffer(&inlineBuffer, matchedRange.location) )) {
                            for (j = matchedRange.location-1; j >= (NSInteger) strRange.location; j--) {
                                if (CFCharacterSetIsCharacterMember(uppercase, CFStringGetCharacterFromInlineBuffer(&inlineBuffer, j) ))
                                    score--;
                                else
                                    score -= SKIPPED_SCORE;
                            }
                        } else {
                            score -= (matchedRange.location-strRange.location)/2;
                        }
                    }
                    score += remainingScore*remainingStrRange.length;
                    score /= strRange.length;
                    CFRelease(userLoc);
                    return score;
                    }
                    }
                    CFRelease(userLoc);
                    return 0;
                    }
                */
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedItem();
        }

        private void MoveSelectedItem()
        {
            //Globals.ThisAddIn.Application.ActiveExplorer().SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);

            var folder = listBox.SelectedItem as FolderWrapper;
            if (folder is null)
            {
                return;
            }

            var explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            for (int i = 1; i <= explorer.Selection.Count; i++)
            {
                var selection = explorer.Selection[i];
                if (selection is Outlook.MailItem)
                {
                    var mailItem = selection as Outlook.MailItem;
                    mailItem.Move(folder.folder);
                }
            }
        }

        private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            int n_item = listBox.Items.Count;
            var i = listBox.SelectedIndex;

            if (n_item == 0)
            {
                return;
            }

            switch (e.Key)
            {
                case Key.Up:
                    if (i > 0)
                    {
                        listBox.SelectedIndex = i - 1;
                    }
                    break;
                case Key.Down:
                    if (i < n_item - 1)
                    {
                        listBox.SelectedIndex = i + 1;
                    }
                    break;
                case Key.Enter:
                    MoveSelectedItem();
                    break;
                default:
                    textBlock.Text += " " + e.Key + "\n";
                    break;
            }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Escape:
                    textBox.Text = "";
                    break;
                default:
                    break;
            }
        }
        public void UpdateFolderList()
        {
            if (folderTree is null)
            {
                Outlook.Folder root = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                // or loop over Application.Session.Stores

                folderTree = new FolderWrapper(root, foldersCollection, null);
                foldersCollection = new ObservableCollection<FolderWrapper>(); //Change to incremental update later *****
                listBox.ItemsSource = foldersCollection;
                ListCollectionView collectionView = (ListCollectionView)CollectionViewSource.GetDefaultView(listBox.ItemsSource);
                collectionView.Filter = FilterHelper;
                collectionView.CustomSort = new SortHelper("");
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

            if (listBox.SelectedIndex < 0)
            {
                listBox.SelectedIndex = 0;
            }
        }

        /*
        private void EnumerateFoldersInDefaultStore(ObservableCollection<FolderWrapper> container)
        {
            Outlook.Folder root =
                Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            // or loop over Application.Session.Stores
            EnumerateFolders(root, container);
        }

        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders(Outlook.Folder folder, ObservableCollection<FolderWrapper> container)
        {
            List<String> ret = new List<String>();
            Outlook.Folders childFolders =
                folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    container.Add(new FolderWrapper(childFolder));

                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder, container);
                }
            }
        }
        */

        /* 
         * Folder Change Events

            Given a collection of folders in Outlook, several events are raised when folders in that collection change:
            Folders.FolderAdd is raised on a Folders collection when a new folder is added. Outlook passes a folder parameter of type MAPIFolder representing the newly added folder.
            Folders.FolderRemove is raised on a Folders collection when a folder is deleted.
            Folders.FolderChange is raised on a Folders collection when a folder is changed. Examples of changes include when the folder is renamed or when the number of items in the folder changes. Outlook passes a folder parameter of type MAPIFolder representing the folder that has changed.
            Listing 10-4 shows an add-in that handles folder change events for any subfolders under the Inbox folder. To get to a Folders collection, we first get a NameSpace object. The NameSpace object is accessed by calling the Application.Session property. The NameSpace object has a method called GetDefaultFolder that returns a MAPIFolder object to which you can pass a member of the enumeration OlDefaultFolders to get a standard Outlook folder. In Listing 10-4, we pass olFolderInbox to get a MAPIFolder for the Inbox. We then connect our event handlers to the Folders collection associated with the Inbox's MAPIFolder object.
            Listing 10-4. A VSTO Add-In That Handles Folder Change Events
            namespace OutlookAddin1
            {
             public partial class ThisApplication
             {
             Outlook.Folders folders;
             private void ThisApplication_Startup(object sender, EventArgs e)
             {
             Outlook.NameSpace ns = this.Session;
             Outlook.MAPIFolder folder = ns.GetDefaultFolder(
             Outlook.OlDefaultFolders.olFolderInbox);
             folders = folder.Folders;

             

             


             }
            }*/
    }

    public class FolderWrapper
    {
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

            (Globals.ThisAddIn.taskPaneControl.taskPaneControl as TaskPaneControl).UpdateFolderList();
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

            (Globals.ThisAddIn.taskPaneControl.taskPaneControl as TaskPaneControl).UpdateFolderList();
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
}
