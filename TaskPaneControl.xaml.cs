using Microsoft.Office.Interop.Outlook;
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

namespace QuickFile
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class TaskPaneControl : UserControl
    {
        private ObservableCollection<String> foldersCollection;

        public TaskPaneControl()
        {
            InitializeComponent();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox senderTb = sender as TextBox;
        
            if (!(listBox is null))
            {
                CollectionViewSource.GetDefaultView(listBox.ItemsSource).Refresh();
            }
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

        private float SortHelper(string str)
        {
            return 0;
            // https://github.com/quicksilver/Quicksilver/blob/8be7395b795179cf51cf30ebf82779e0f9ba2138/Quicksilver/Code-QuickStepFoundation/QSSense.m
            /* 
             * 
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
        var folder_names = EnumerateFoldersInDefaultStore();
        //folder_names.ForEach(s => foldersCollection.Add(s));
        foldersCollection = new ObservableCollection<string>(folder_names); //Change to incremental update later *****
        listBox.ItemsSource = foldersCollection;

        CollectionView collectionView = (CollectionView)CollectionViewSource.GetDefaultView(listBox.ItemsSource);
        collectionView.Filter = FilterHelper;
            listBox.SelectedIndex = 0;
        }

        private List<String> EnumerateFoldersInDefaultStore()
        {
        Outlook.Folder root =
            Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
        return EnumerateFolders(root);
        }

        // Uses recursion to enumerate Outlook subfolders.
        private List<String> EnumerateFolders(Outlook.Folder folder, String prefix="")
        {
        List<String> ret = new List<String>(); 
        Outlook.Folders childFolders =
            folder.Folders;
        if (childFolders.Count > 0)
        {
            foreach (Outlook.Folder childFolder in childFolders)
            {
                // Write the folder path.
                ret.Add(childFolder.FolderPath);
                // Call EnumerateFolders using childFolder.
                ret.AddRange(EnumerateFolders(childFolder));
            }
        }
        return ret;
        }
        }
}
