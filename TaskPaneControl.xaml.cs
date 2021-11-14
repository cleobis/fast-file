using ControlzEx.Standard;
using MahApps.Metro.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Xaml.Behaviors.Core;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
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

        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        private ThisAddIn addIn = null;
        private ListCollectionView listCollectionView = null;
        internal TaskPaneContext taskPaneContext = null;

        public TaskPaneControl()
        {
            this.InheritanceBehavior = InheritanceBehavior.SkipToAppNext;
            InitializeComponent();
        }


        public void SetUp()
        {
            addIn = Globals.ThisAddIn;
        }

        public void LateSetup()
        {
            // Now that plugin initializatin code is moved out of Startup into an asynchronous code, need to make sure it is done before trying to reference. Should really have some feedback on the UI.
            if (listCollectionView == null) {
                listCollectionView = new ListCollectionView(addIn.foldersCollection);
                var sortHelper = new SortHelper(textBox);
                listCollectionView.Filter = o => sortHelper.Filter(o);
                listCollectionView.CustomSort = sortHelper;
                listBox.ItemsSource = listCollectionView;
            }
        }

        internal void RefreshSelection(FolderWrapper folder = null)
        {
            //listCollectionView.Refresh();
            this.Dispatcher.Invoke(() => // Must be on main thread
            {
                if (folder != null)
                {
                    listBox.SelectedItem = folder;
                }

                if (listBox.SelectedIndex < 0)
                {
                    listBox.SelectedIndex = 0;
                }
            });
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
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
            catch (System.Exception err)
            {
                MessageBox.Show("Unexpected error processing TextBox_TextChanged.\n" + err.Message, "Fast File Error");
                Logger.Error(err, "Unexpected error processing TextBox_TextChanged.");
            }
        } 

        internal class SortHelper : System.Collections.IComparer
        {
            private static CharacterSet WordSeperators = null;
            private static CharacterSet Uppercase = null;
            private readonly TextBox textBox;
            private string textBoxCache;
            private Dictionary<String, float> scoreCache;

            public SortHelper(TextBox _textBox)
            {
                textBox = _textBox;
                textBoxCache = textBox.Text;

                if (WordSeperators == null)
                {
                    // Initialize static member once
                    WordSeperators = new CharacterSet();
                    WordSeperators.Add(UnicodeCategory.SpaceSeparator);
                    WordSeperators.Add('.');
                    WordSeperators.Add('\\');
                }
                if (Uppercase == null)
                {
                    Uppercase = new CharacterSet();
                    Uppercase.Add(UnicodeCategory.UppercaseLetter);
                }

                scoreCache = new Dictionary<String, float>();
            }

            private struct StringRange
            {
                public int Start, Length;
                public StringRange(int _start, int _length)
                {
                    Start = _start;
                    Length = _length;
                }

                public String Slice(String str)
                {
                    return str.Substring(Start, Length);
                }
            };
            private class CharacterSet
            {
                private List<Char> _characters;
                private List<UnicodeCategory> _categories;
                public CharacterSet()
                {
                    _characters = new List<Char>();
                    _categories = new List<UnicodeCategory>();
                }
                public void Add(char _char)
                {
                    _characters.Add(_char);
                }
                public void Add(UnicodeCategory category)
                {
                    _categories.Add(category);
                }
                public bool Contains(char _char)
                {
                    if (_characters.Contains(_char))
                    {
                        return true;
                    }

                    return _categories.Contains(Char.GetUnicodeCategory(_char));
                }
            }

            public bool Filter(object obj)
            {
                CheckTextBoxUnchanged();

                if (string.IsNullOrEmpty(textBoxCache))
                {
                    return true;
                }
                else
                {
                    var fw = obj as FolderWrapper;
                    float score = GetScore(fw.folder.Name);
                    return score > 0;
                }
            }

            public int Compare(object a, object b)
            {
                CheckTextBoxUnchanged();

                var fwa = a as FolderWrapper;
                var fwb = b as FolderWrapper;

                if (string.IsNullOrEmpty(textBoxCache))
                {
                    return String.Compare(fwa.folder.FolderPath, fwb.folder.FolderPath, true);
                }
                else
                {
                    // return -ve if a < b logically. 
                    float score_a = GetScore(fwa.folder.Name);
                    float score_b = GetScore(fwb.folder.Name);
                    float diff = score_a - score_b;
                    if (score_a > score_b)
                        return -1;
                    else if (score_a == score_b)
                        return 0;
                    else
                        return +1;
                }
            }

            private float GetScore(string input)
            {
                if (string.IsNullOrEmpty(input))
                    return 0;

                float finalScore;
                bool cacheHit = scoreCache.TryGetValue(input, out finalScore);
                if (!cacheHit)
                {
                    finalScore = ScoreForAbbreviation(input, textBoxCache);
                    scoreCache.Add(input, finalScore);
                }
                return finalScore;

                // https://github.com/quicksilver/Quicksilver/blob/8be7395b795179cf51cf30ebf82779e0f9ba2138/Quicksilver/Code-QuickStepFoundation/QSSense.m
                float ScoreForAbbreviation(String str, String abbr)
                {
                    return ScoreForAbbreviationWithRanges(str, abbr, new StringRange(0, str.Length), new StringRange(0, abbr.Length));
                }

                float ScoreForAbbreviationWithRanges(String str, String abbr, StringRange strRange, StringRange abbrRange)
                {
                    const float IGNORED_SCORE = 0.9f;
                    const float SKIPPED_SCORE = 0.15f;

                    if (abbrRange.Length == 0)
                        return IGNORED_SCORE; //deduct some points for all remaining letters

                    if (abbrRange.Length > strRange.Length)
                        return 0.0f;

                    float score = 0.0f, remainingScore = 0.0f;
                    int i, j;
                    StringRange matchedRange, remainingStrRange = new StringRange(0, 0), adjustedStrRange = strRange;

                    for (i = abbrRange.Length; i > 0; i--)
                    {
                        //Search for steadily smaller portions of the abbreviation
                        String curAbbr = abbr.Substring(abbrRange.Start, i);
                        int idx = str.IndexOf(curAbbr, adjustedStrRange.Start, adjustedStrRange.Length - abbrRange.Length + i, StringComparison.CurrentCultureIgnoreCase);
                        matchedRange = new StringRange(idx, curAbbr.Length);
                        if (idx == -1)
                        {
                            // not found
                            continue;
                        }

                        remainingStrRange.Start = matchedRange.Start + matchedRange.Length;
                        remainingStrRange.Length = strRange.Start + strRange.Length - remainingStrRange.Start;

                        // Search what is left of the string with the rest of the abbreviation
                        remainingScore = ScoreForAbbreviationWithRanges(str, abbr, remainingStrRange, new StringRange(abbrRange.Start + i, abbrRange.Length - i));

                        if (remainingScore != 0)
                        {
                            score = remainingStrRange.Start - strRange.Start;
                            // ignore skipped characters if is first letter of a word
                            if (matchedRange.Start > strRange.Start)
                            {//if some letters were skipped

                                if (WordSeperators.Contains(str.ElementAt(matchedRange.Start - 1)))
                                {
                                    for (j = matchedRange.Start - 2; j >= strRange.Start; j--)
                                    {
                                        if (WordSeperators.Contains(str.ElementAt(j)))
                                            score--;
                                        else
                                            score -= SKIPPED_SCORE;
                                    }
                                }
                                else if (Uppercase.Contains(str.ElementAt(matchedRange.Start)))
                                {
                                    for (j = matchedRange.Start - 1; j >= strRange.Start; j--)
                                    {
                                        if (Uppercase.Contains(str.ElementAt(j)))
                                            score--;
                                        else
                                            score -= SKIPPED_SCORE;
                                    }
                                }
                                else
                                {
                                    score -= (matchedRange.Start - strRange.Start) / 2;
                                }
                            }
                            score += remainingScore * remainingStrRange.Length;
                            score /= strRange.Length;
                            return score;
                        }
                    }
                    return 0.0f;
                }

            }

            private void CheckTextBoxUnchanged()
            {
                if (textBox.Text == textBoxCache)
                    return;
                scoreCache.Clear();
                textBoxCache = textBox.Text;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MoveSelectedItem();
            }
            catch (System.Exception err)
            {
                MessageBox.Show("Unexpected error processing burron.\n" + err.Message, "Fast File Error");
                Logger.Error(err, "Unexpected error processing button.");
            }
        }

        private void MoveSelectedItem()
        {
            var folder = listBox.SelectedItem as FolderWrapper;
            if (folder is null)
            {
                return;
            }
            taskPaneContext.MoveSelectedItem(folder.folder);
            taskPaneContext.Visible = false;
        }

        private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            int n_item = listBox.Items.Count;
            var i = listBox.SelectedIndex;

            Logger.Debug("TextBox PreviewKeydown {key}.", e.Key);

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
                    e.Handled = true;
                    break;
                case Key.Down:
                    if (i < n_item - 1)
                    {
                        listBox.SelectedIndex = i + 1;
                    }
                    e.Handled = true;
                    break;
                /*case Key.Enter:
                    MoveSelectedItem();
                    break;*/
                default:
                    break;
            }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            Logger.Debug("TextBox Keydown {key}", e.Key);
            switch (e.Key)
            {
                case Key.Escape:
                    if (textBox.Text != "")
                    {
                        textBox.Text = "";
                        e.Handled = true;
                    }
                    else
                    {
                        //taskPaneContext.Visible = false;
                    }
                    break;
                default:
                    break;
            }
        }

        private void UserControl_KeyDown(object sender, KeyEventArgs e)
        {
            Logger.Debug("UserControl_KeyDown {key}", e.Key);
            switch (e.Key)
            {
                case Key.Escape:
                    taskPaneContext.Visible = false;
                    e.Handled = true;
                    break;
                case Key.Enter:
                    MoveSelectedItem();
                    e.Handled = true;
                    break;
                default:
                    break;
            }
        }
        private void listBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                MoveSelectedItem();
            }
            catch (System.Exception err)
            {
                MessageBox.Show("Unexpected error double click.\n" + err.Message, "Fast File Error");
                Logger.Error(err, "Unexpected error double click.");
            }
        }
        
        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            Logger.Debug("UserControl_PreviewKeyDown {key}", e.Key);
        }

        private void textBox_GotFocus(object sender, RoutedEventArgs e)
        {
            Logger.Debug("textBox_GotFocus");
        }

        private void textBox_LostFocus(object sender, RoutedEventArgs e)
        {
            Logger.Debug("textBox_LostFocus");
        }

        private void textBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            Logger.Debug("textBox_GotKeyboardFocus");
        }

        private void textBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            Logger.Debug("textBox_LostKeyboardFocus");
        }
    }
    public class DesignerMockData
    {
        public DesignerMockData()
        {
        }
        public String DisplayName
        {
            get { return "DisplayName"; }
            set { }
        }
        public String DisplayPath
        {
            get { return "DisplayPath"; }
            set { }
        }

        public Thickness DisplayNameMargin
        {
            get { return new Thickness(20, 0, 0, 0); }
            set { }
        }
    }
    
}
