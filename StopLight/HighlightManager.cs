using System;
using System.Collections.Generic;
using System.Linq;
using Forms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace StopLight
{
    class HighlightManager
    {
        //store data we read in from the Excel file
        private Dictionary<string, int> Dictionary = new Dictionary<string, int>();
        private List<string> IndexNames = new List<string>();
        //save already highlighted words so we can re-highlight them when the user resets the highlight
        private List<Tuple<Word.Range, Word.WdColorIndex>> SavedHighlights =
            new List<Tuple<Word.Range, Word.WdColorIndex>>();
        //store unknown words
        private HashSet<string> UnknownWords = new HashSet<string>();
        //save range of previous highlight
        private Word.Range prevRange = null;
        //this helps prevent duplicate words (for example, run and running, horse and horses)
        private LemmatizerImplementation lemmatizer = new LemmatizerImplementation();
        //helper class for connecting to excel file
        private OleDbHelper oleDbHelper = new OleDbHelper();
        private bool SafeToContinueHighlight()
        {
            //checks
            //1. selects are set
            //2. if something is already highlighted,
            // make sure it's ok with the user

            //first see if file is set
            if (!this.FileSelected())
            {
                Forms.MessageBox.Show(Strings.fileNotSelectedMessage, Strings.fileNotSelectedCaption);
                return false;
            }

            //validate selects
            if (!Globals.ThisAddIn.Ribbon.SelectsSet())
            {
                Forms.MessageBox.Show(Strings.rangeNotSetMessage, Strings.rangeNotSetCaption);
                return false;
            }

            //Check here if any part of the text is already highlighted
            //If so, warn the user
            Word.Range range = null;
            if (prevRange != null)
                range = prevRange;
            else
                range = this.FindRange();

            if (range.HighlightColorIndex != Colors.none)
            {
                //format the message
                Forms.MessageBoxButtons buttons = Forms.MessageBoxButtons.YesNo;
                Forms.DialogResult result;
                // Displays the warning
                result = Forms.MessageBox.Show(Strings.alreadyHighlightedMessage, Strings.alreadyHighlightedCaption, buttons);

                //can be abort or no, so just say 
                //if not yes then abort operation
                if (result != Forms.DialogResult.Yes)
                    return false;
            }

            return true;
        }

        internal void Highlight()
        {
            if (!this.SafeToContinueHighlight())
                return;

            Word.Range range = this.FindRange();
            this.UndoHighlightToOriginal();
            UnknownWords.Clear();

            //it's logical to save the existing highlights
            //before highlighting but
            //this will increase run time by 2n
            //so actually do this ↓ while running loop to highlight
            //this.SaveExistingHighlight(range);

            //find numerical indexes of the highlight range
            LowerAndUpperBounds bounds = new LowerAndUpperBounds(Globals.ThisAddIn.Ribbon, this.GetIndexNames());
            bounds.FindLowerUpperBounds();

            //this loop also saves highlighted text
            this.SaveExistingAndHighlightRange(range, bounds);

        }



        //note that this saves existing highlight and highlights
        //this is because looping through document twice will increase runtime by 2n
        //where n is around 15 seconds per page \(*o*)/
        private void SaveExistingAndHighlightRange(Word.Range range, LowerAndUpperBounds bounds)
        {
            //range to search
            int wordCount = range.Words.Count;
            //***lower bound for arrays is 1 in Word object model***
            // (array[0] does exist but is not used)
            // (array[1] is the first word in the array)
            int lowerBoundWordCount = 1;
            int upperBoundWordCount = wordCount + 1;
            Word.Words words = range.Words;
            //we have two saved colors (one for continuing range of same color)
            //so we can bulk highlight a sequence of words with the same color
            //for faster runtime (updating word one by one takes too long)
            Word.WdColorIndex tempColor = Word.WdColorIndex.wdNoHighlight;
            Word.WdColorIndex tempColor2 = Word.WdColorIndex.wdNoHighlight;
            Word.Range tempRange = null;

            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorWait;

            for (int i = lowerBoundWordCount; i < upperBoundWordCount; i++)
            {
                Word.Range wordRange = words[i];
                //save existing highlight first
                if (wordRange.HighlightColorIndex != Colors.none)
                    this.SaveExistingHighlight(wordRange);

                //get word
                string word = wordRange.Text;
                string cleanWord = CleanWord(word);

                //move page as necessary (kinda like progress bar but more intuitive)
                //right now, every 20 words
                if (i % 20 == 0)
                    wordRange.Select();


                /*
                Periods, commas, etc are considered words
                so skip through them if so.
                However, for hyphens it may be one word
                for example merry-go-round
                or it may not be like -this hyphen acts as a bullet point.
                Make an exception for hyphens and consider it as one word if necessary
                */
                if (isPunctuation(cleanWord))
                {
                    if (tempRange == null)
                        continue;
                    //reset highlight color to nothing
                    this.AddColoring(tempRange, tempColor);
                    tempRange = null;
                    tempColor = Word.WdColorIndex.wdNoHighlight;
                    wordRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                    continue;
                }

                //now highlight

                /*
                For highlighting, looping word by word is extremely slow
                900 words about 1 minute
                But highlighting the entire page a single color takes a second..
                So to lessen run time, bulk highlight words with same color.

                */

                if (Dictionary.ContainsKey(cleanWord))
                {
                    int wordIndex = Dictionary[cleanWord];
                    tempColor2 = bounds.GetColorOfIndex(wordIndex);

                }
                else
                {
                    tempColor2 = Colors.unknown;
                    UnknownWords.Add(cleanWord);
                }

                if (tempRange == null) //first instance
                {
                    tempColor = tempColor2;
                    tempRange = wordRange;
                }
                else
                {
                    if (tempColor == tempColor2) //we continue to the next word
                    {
                        int rangeSize = wordRange.Characters.Count;
                        tempRange.MoveEnd(Word.WdUnits.wdCharacter, rangeSize);
                    }
                    else //we highlight the range before
                    {
                        this.AddColoring(tempRange, tempColor);

                        tempRange = wordRange;
                        tempColor = tempColor2;
                    }
                }


            } //end for loop of words
            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorNormal; //pointer back to normal
            //highlight last set of words
            if (tempRange != null) //in case last word is a punctuation mark
                this.AddColoring(tempRange, tempColor);
            //remove selected text
            Globals.ThisAddIn.Application.Selection.Move();
            //save range so the next time we highlight,
            //we can remove highlight on this range
            prevRange = range;
        }

        private void AddColoring(Word.Range range, Word.WdColorIndex color)
        {
            range.HighlightColorIndex = color;
        }

        //indicates which words highlight() should skip over
        private bool isPunctuation(string word)
        {
            if (word.All(char.IsPunctuation))
                return true;
            else
                return false;
        }

        private void SaveExistingHighlight(Word.Range range)
        {
            //go through every character
            //since the whole word might not be highlighted
            Word.Characters characters = range.Characters;
            int chCt = characters.Count;
            int chCtUpperBound = chCt + 1;
            int chCtLowerBound = 1;
            for (int i = chCtLowerBound; i < chCtUpperBound; i++)
            {
                Word.Range r = characters[i];
                if (r.HighlightColorIndex != Colors.none)
                {
                    SavedHighlights.Add(Tuple.Create(r, r.HighlightColorIndex));
                }
            }
        }
        

        //set highlights back to original
        private void ReHighlight()
        {
            List<int> tempArray = new List<int>();
            for (int i = 0; i < SavedHighlights.Count; i++)
            {
                Tuple<Word.Range, Word.WdColorIndex> tuple = SavedHighlights[i];
                Word.Range tRange = tuple.Item1;
                Word.WdColorIndex tHighlight = tuple.Item2;

                //we should clear the range's highlight before calling this function
                if (tRange.HighlightColorIndex == Colors.none) //just in case
                {
                    tRange.HighlightColorIndex = tHighlight;
                    tempArray.Add(i);
                }
            }

            //loop backwards so we can remove elements as we go
            int tempCt = tempArray.Count - 1;
            for (int i = tempCt; i >= 0; i--)
            {
                SavedHighlights.RemoveAt(tempArray[i]);
            }
        }
        
        private Word.Range FindRange()
        {
            //we need to see if we are looking at the whole document or a selected area.
            //If the user selected an area, only that part should be checked for highlighting.
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            Word.Range range = null;
            if (selection.Start == selection.End) //no selection
            {
                range = Globals.ThisAddIn.Application.ActiveDocument.Content; //whole document
            }
            else
            {
                range = selection.Range;
            }

            return range;
        }

        internal string CleanWord(string word)
        {
            if (word == null)
                return "";

            word = word.Trim();
            word = word.ToLower();

            //use an open-source lemmatizer to make words their root form
            //ie is    -> be
            //   hills -> hill
            word = lemmatizer.Lemmatize(word);

            return word;
        }



        internal void UndoHighlightToOriginal()
        {
            if (prevRange == null)
                return;
            //clear styling
            prevRange.HighlightColorIndex = Colors.none;

            //put original styling back
            this.ReHighlight();
            //empty the list
            //we are doing this in rehighlight.
            //note that rehighlight may still leave some characters in the list
            //SavedHighlights.Clear();
            prevRange = null;
        }

        //IList is a readonly list
        internal IList<string> GetIndexNames()
        {
            return IndexNames.AsReadOnly();
        }

        internal IList<string> GetUnknownWords()
        {
            return UnknownWords.ToList().AsReadOnly();
        }

        internal void ClearData()
        {
            Dictionary.Clear();
            IndexNames.Clear();
            UnknownWords.Clear();
        }


        internal bool FileSelected()
        {
            if (Dictionary.Count == 0)
                return false;
            if (IndexNames.Count == 0)
                return false;
            return true;
        }

        internal void ProcessFile(string fileName)
        {
            //clear all existing data before starting
            this.ClearData();
            oleDbHelper.ProcessFile(fileName, ref IndexNames, ref Dictionary);

        }

        internal bool CanAddToDictionary(string word)
        {
            if (Dictionary.ContainsKey(word))
                return false;
            else
                return true;
        }

        internal void AddWordToExcel(string fileName, string word, string columnName)
        {
            oleDbHelper.AddWordToExcel(fileName, word, columnName);
            this.UpdateUnknownWord(word, columnName);
        }

        private void UpdateUnknownWord(string word, string columnName)
        {
            //we need to add the word to the dictionary and
            //remove it from the unknown words list
            int columnIndex = IndexNames.IndexOf(columnName);
            if (!Dictionary.ContainsKey(word))
                Dictionary.Add(word, columnIndex);


            UnknownWords.Remove(word);
        }

        internal void SelectUnknownWordsToAdd(string fileName)
        {
            if (!this.FileSelected())
            {
                Forms.MessageBox.Show(Strings.fileNotSelectedMessage, Strings.fileNotSelectedCaption);
                return;
            }
            new SelectWordToAddForm(fileName).ShowDialog();
        }

        internal void AddUnknownWord(string fileName, string word, SelectWordToAddForm form = null)
        {
            if (!this.FileSelected())
            {
                Forms.MessageBox.Show(Strings.fileNotSelectedMessage, Strings.fileNotSelectedCaption);
                return;
            }

            if (!this.CanAddToDictionary(word))
            {
                Forms.MessageBox.Show(Strings.duplicateWordText, Strings.duplicateWordCaption);
                return;

            }

            //ask what column to put it in
            new AddWordForm(word, fileName, form).ShowDialog();

        }
    }
}
