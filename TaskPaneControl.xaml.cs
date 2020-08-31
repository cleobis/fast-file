using ControlzEx.Standard;
using MahApps.Metro.Controls;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Xaml.Behaviors.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
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
// * Improve styling
// * Guess item from conversation
// * Toolbar dropdown?

namespace QuickFile
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class TaskPaneControl : UserControl
    {

        private ThisAddIn addIn = null;
        private ListCollectionView listCollectionView = null;
        internal TaskPaneContext taskPaneContext = null;
        
        public TaskPaneControl()
        {
            InitializeComponent();
        }

        
        public void SetUp()
        {
            addIn = Globals.ThisAddIn;

            listCollectionView = new ListCollectionView(addIn.foldersCollection);
            listCollectionView.Filter = FilterHelper;
            listCollectionView.CustomSort = new SortHelper("");
            listBox.ItemsSource = listCollectionView;
        }
        
        internal void RefreshSelection()
        {
            if (listBox.SelectedIndex < 0)
            {
                listBox.SelectedIndex = 0;
            }
        }
        
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox senderTb = sender as TextBox;

            if (!(listBox is null))
            {
                listCollectionView.Refresh();
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
            var folder = listBox.SelectedItem as FolderWrapper;
            if (folder is null)
            {
                return;
            }
            taskPaneContext.MoveSelectedItem(folder.folder);
        }

        private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            int n_item = listBox.Items.Count;
            var i = listBox.SelectedIndex;

            Debug.WriteLine("TextBox PreviewKeydown " + e.Key);

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
                    break;
            }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            Debug.WriteLine("TextBox Keydown " + e.Key);
            switch (e.Key)
            {
                case Key.Escape:
                    if (textBox.Text != "")
                    {
                        textBox.Text = "";
                    }
                    else
                    {
                        taskPaneContext.Visible = false;
                    }
                    break;
                default:
                    break;
            }
        }

        private void UserControl_KeyDown(object sender, KeyEventArgs e)
        {
            Debug.WriteLine("UserControl_KeyDown " + e.Key);
        }

        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            Debug.WriteLine("UserControl_PreviewKeyDown " + e.Key);
        }

        private void textBox_GotFocus(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("textBox_GotFocus");
        }

        private void textBox_LostFocus(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("textBox_LostFocus");
        }

        private void textBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            Debug.WriteLine("textBox_GotKeyboardFocus");
        }

        private void textBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            Debug.WriteLine("textBox_LostKeyboardFocus");
        }
    }
}
