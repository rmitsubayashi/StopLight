using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Forms = System.Windows.Forms;

namespace StopLight
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        
        private String fileName = "";
        
        private String _through = " ~ ";


        //how many select boxes there are on the ribbon (red, yellow, green, none)
        readonly private int totalSelectCt = 4;
        //select options for teh select boxes
        private List<string> noneDropdownLabels = new List<string>();
        private List<string> greenDropdownLabels = new List<string>();
        private List<string> yellowDropdownLabels = new List<string>();
        //we mark selected item so we can grab it easily when highlighting
        private int noneDropdownIndex = 0;
        private int greenDropdownIndex = 0;
        private int yellowDropdownIndex = 0;
        //enable and disable selects
        private bool noneSelectEnabled = false;
        private bool greenSelectEnabled = false;
        private bool yellowSelectEnabled = false;
        //for range labels
        private String noneLowerLabel = "";
        private String greenLowerLabel = "";
        private String yellowLowerLabel = "";
        private String redLowerLabel = "";
        private String redUpperLabel = "";

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("StopLight.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        //this is for refreshing all labels
        //need it to update after language change / file name change
        public string ActionsGroupLabel(Office.IRibbonControl control)
        {
            return Strings.actionsGroup;
        }

        public string HighlightButtonLabel(Office.IRibbonControl control)
        {
            return Strings.highlightButton;
        }

        public string RemoveButtonLabel(Office.IRibbonControl control)
        {
            return Strings.removeButton;
        }

        public string SelectFileGroupLabel(Office.IRibbonControl control)
        {
            return Strings.selectFileGroup;
        }

        public string SelectFileButtonLabel(Office.IRibbonControl control)
        {
            return Strings.selectFileButton;
        }

        public string FileNameLabelLabel(Office.IRibbonControl control)
        {
            return Strings.fileNameLabel;
        }

        public string DropdownNoneLabel(Office.IRibbonControl control)
        {
            return Strings.dropdownNoneLabel;
        }

        public string DropdownGreenLabel(Office.IRibbonControl control)
        {
            return Strings.dropdownGreenLabel;
        }

        public string DropdownYellowLabel(Office.IRibbonControl control)
        {
            return Strings.dropdownYellowLabel;
        }

        public string DropdownRedLabel(Office.IRibbonControl control)
        {
            return Strings.dropdownRedLabel;
        }

        public string RangeLabel(Office.IRibbonControl control)
        {
            return Strings.rangeGroup;
        }

        public string AddUnknownWordsGroupLabel(Office.IRibbonControl control)
        {
            return Strings.addUnknownWordsGroup;
        }

        public string AddUnknownWordsButtonLabel(Office.IRibbonControl control)
        {
            return Strings.addUnknownWordsButton;
        }

        //enabled/disabled
        public bool SelectEnabled(Office.IRibbonControl control)
        {
            if (control.Id.Equals("Dropdown_None_Upper_Select"))
                return noneSelectEnabled;
            else if (control.Id.Equals("Dropdown_Green_Upper_Select"))
                return greenSelectEnabled;
            else if (control.Id.Equals("Dropdown_Yellow_Upper_Select"))
                return yellowSelectEnabled;
            else return false;
        }

        //refresh dropdown labels after selecting file
        public string UpdateLowerLabel(Office.IRibbonControl control)
        {
            string returnLabel = "";
            if (control.Id.Equals("Dropdown_None_Lower_Label"))
            {
                if (noneLowerLabel.Equals(""))
                    returnLabel = Strings.lowerSelectDefault;
                else
                    returnLabel = noneLowerLabel;
            }
            else if (control.Id.Equals("Dropdown_Green_Lower_Label"))
            {
                if (greenLowerLabel.Equals(""))
                    returnLabel = Strings.lowerSelectDefault;
                else
                    returnLabel = greenLowerLabel;
            }
            else if (control.Id.Equals("Dropdown_Yellow_Lower_Label"))
            {
                if (yellowLowerLabel.Equals(""))
                    returnLabel = Strings.lowerSelectDefault;
                else
                    returnLabel = yellowLowerLabel;
            }
            else if (control.Id.Equals("Dropdown_Red_Lower_Label"))
            {
                if (redLowerLabel.Equals(""))
                    returnLabel = Strings.lowerSelectDefault;
                else
                    returnLabel = redLowerLabel;
            }

            return returnLabel + _through;

        }

        public string UpdateRedUpperLabel(Office.IRibbonControl control)
        {
            if (redUpperLabel.Equals(""))
                return Strings.lowerSelectDefault;
            else
            {
                return redUpperLabel;
            }
        }

        public int DropdownCount(Office.IRibbonControl control)
        {
            if (control.Id.Equals("Dropdown_None_Upper_Select"))
                return noneDropdownLabels.Count;
            else if (control.Id.Equals("Dropdown_Green_Upper_Select"))
                return greenDropdownLabels.Count;
            else if (control.Id.Equals("Dropdown_Yellow_Upper_Select"))
                return yellowDropdownLabels.Count;
            else return 0;
        }

        public String DropdownItemLabels(Office.IRibbonControl control, int index)
        {
            if (control.Id.Equals("Dropdown_None_Upper_Select"))
                return noneDropdownLabels[index];
            else if (control.Id.Equals("Dropdown_Green_Upper_Select"))
                return greenDropdownLabels[index];
            else if (control.Id.Equals("Dropdown_Yellow_Upper_Select"))
                return yellowDropdownLabels[index];
            else return "";
        }

        public int SelectIndex(Office.IRibbonControl control)
        {
            if (control.Id.Equals("Dropdown_None_Upper_Select"))
                return noneDropdownIndex;
            else if (control.Id.Equals("Dropdown_Green_Upper_Select"))
                return greenDropdownIndex;
            else if (control.Id.Equals("Dropdown_Yellow_Upper_Select"))
                return yellowDropdownIndex;
            else return 0;
        }

        //actions for ribbon buttons
        
        public void Highlight(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.HighlightManager.Highlight();
        }

        public void RemoveHighlight(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.HighlightManager.UndoHighlightToOriginal();
        }

        public void SelectFile(Office.IRibbonControl control)
        {
            this.SelectFile();

        }

        public void SelectChanged(Office.IRibbonControl control, string id, int index)
        {
            bool changeAllowed = this.SelectChangeAllowed(id, index);

            if (changeAllowed)
            {
                this.ChangeSelect(id, index);
            } else
            {
                //reset back to state before attempted change
                this.RenderDropdowns();
            }
        }

        public void AddUnknownWords(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.HighlightManager.SelectUnknownWordsToAdd(this.fileName);
        }

        //implementations of actions

        private void SelectFile()
        {
            // 1. Allow user to select Excel file.
            // 2. Validate file.
            // 3. Read file and save into Word Document (not Ribbon)
            // 4. Update ranges on the Ribbon

            //set form
            Forms.OpenFileDialog openFileDialog = new Forms.OpenFileDialog();
            //initial directory should be the directory where we store our example files
            //CHANGE!!!
            openFileDialog.InitialDirectory = "/";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //this sets the directory of the computer to the state before we opened this
            //since we are not requiring the user to open the same directory multiple times,
            //it is more user friendly to set back to original directory
            openFileDialog.RestoreDirectory = true;

            Forms.DialogResult result;
            result = openFileDialog.ShowDialog();

            if (result == Forms.DialogResult.OK)
            {
                //first check if the file exists
                try
                {
                    if (openFileDialog.CheckFileExists)
                    {
                        string fileName = openFileDialog.FileName;
                        Globals.ThisAddIn.HighlightManager.ProcessFile(fileName);

                        //Update ribbon
                        UpdateRibbonFileName(fileName);
                        UpdateSelects();
                    }
                }
                catch (ArgumentNullException)
                {
                    Forms.MessageBox.Show(Strings.cannotReadFileMessage, Strings.cannotReadFileCaption);
                }
            }
        }

        private void UpdateRibbonFileName(string fileName)
        {
            string fileNameText = Strings.fileNameLabel;
            string fileLabel = (fileNameText.Split(':'))[0];
            fileLabel += ": ";
            //fileName is currently the whole path
            //so trim down to the actual file name
            int lastIndexFileName = fileName.LastIndexOf('\\');
            string displayFileName = fileName.Substring(lastIndexFileName + 1);
            fileLabel += displayFileName;

            Strings.fileNameLabel = fileLabel;

            //render
            ribbon.InvalidateControl("FileNameLabel");

            //save file name
            //we need it when accessing file again to add unknown words
            this.fileName = fileName;
        }

        

        private void ClearSelects()
        {
            noneDropdownLabels.Clear();
            greenDropdownLabels.Clear();
            yellowDropdownLabels.Clear();
            noneLowerLabel = "";
            greenLowerLabel = "";
            yellowLowerLabel = "";
            redLowerLabel = "";
            redUpperLabel = "";
            noneDropdownIndex = 0;
            greenDropdownIndex = 0;
            yellowDropdownIndex = 0;
            
            noneSelectEnabled = false;
            greenSelectEnabled = false;
            yellowSelectEnabled = false;

            RenderDropdowns();
        }

        private void RenderDropdowns()
        {
            ribbon.InvalidateControl("Dropdown_None_Upper_Select");
            ribbon.InvalidateControl("Dropdown_Green_Upper_Select");
            ribbon.InvalidateControl("Dropdown_Yellow_Upper_Select");
            ribbon.InvalidateControl("Dropdown_None_Lower_Label");
            ribbon.InvalidateControl("Dropdown_Green_Lower_Label");
            ribbon.InvalidateControl("Dropdown_Yellow_Lower_Label");
            ribbon.InvalidateControl("Dropdown_Red_Lower_Label");
            ribbon.InvalidateControl("Dropdown_Red_Upper_Label");
        }

        private void UpdateSelects()
        {
            if (!Globals.ThisAddIn.HighlightManager.FileSelected())
                return;
            
            //reset selects
            ClearSelects();

            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            int columnCt = indexNames.Count();
            //if there aren't enough items on the list to make none red yellow and green
            //one item : green
            //two items : green + red
            //three items : green + yellow + red
            if (columnCt == 1)
                UpdateSelectsOneColumn();
            else if (columnCt == 2)
                UpdateSelectsTwoColumns();
            else if (columnCt == 3)
                UpdateSelectsThreeColumns();
            else //enough columns to fill all four
                UpdateSelectsFourColumns();
            
        }

        private void UpdateSelectsOneColumn()
        {
            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            greenSelectEnabled = true;

            //set start and end
            greenLowerLabel = indexNames[0];
            greenDropdownLabels.Add(indexNames[0]);

            greenDropdownIndex = 0;

            //this is to make sure the dropdown is empty
            //didn't work when I ran this in ClearSelects()
            noneDropdownLabels.Add(" ");
            yellowDropdownLabels.Add(" ");
            RenderDropdowns();
        }

        private void UpdateSelectsTwoColumns()
        {
            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            greenSelectEnabled = true;
            //set start and end
            greenLowerLabel = indexNames[0];
            greenDropdownLabels.Add(indexNames[0]);
            redLowerLabel = indexNames[1];
            redUpperLabel = indexNames[1];

            greenDropdownIndex = 0;

            noneDropdownLabels.Add(" ");
            yellowDropdownLabels.Add(" ");
            RenderDropdowns();
        }

        private void UpdateSelectsThreeColumns()
        {
            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            greenSelectEnabled = true;
            yellowSelectEnabled = true;
            //set start and end
            greenLowerLabel = indexNames[0];
            greenDropdownLabels.Add(indexNames[0]);
            yellowLowerLabel = indexNames[1];
            yellowDropdownLabels.Add(indexNames[1]);
            redLowerLabel = indexNames[2];
            redUpperLabel = indexNames[2];

            greenDropdownIndex = 0;
            yellowDropdownIndex = 0;

            noneDropdownLabels.Add(" ");
            RenderDropdowns();
        }

        private void UpdateSelectsFourColumns()
        {
            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            int columnCt = indexNames.Count();
            noneSelectEnabled = true;
            greenSelectEnabled = true;
            yellowSelectEnabled = true;

            List<string>[] dropdownsToFill =
            {
                    noneDropdownLabels,
                    greenDropdownLabels,
                    yellowDropdownLabels
            };

            //set start and end
            noneLowerLabel = indexNames[0];
            redUpperLabel = indexNames[columnCt - 1];
            //we do not need to loop through last red select (no select box)
            int loopCt = totalSelectCt - 1;
            for (int i = 0; i < loopCt; i++)
            {
                dropdownsToFill[i].Add(Strings.upperSelectDefault);

                int upperBound = columnCt - totalSelectCt + i;

                for (int j = i; j <= upperBound; j++)
                    dropdownsToFill[i].Add(indexNames[j]);
            }

            noneDropdownIndex = 0;
            greenDropdownIndex = 0;
            yellowDropdownIndex = 0;
            RenderDropdowns();
        }


        private bool SelectChangeAllowed(string id, int index)
        {
            //note that this will never be called by < 4 columns

            //the dropdown value has to be 
            // 1. more than the associated label (lowerbound)
            // 2. less than the next dropdown (upperbound)
            //find all the labels/dropdown strings to compare with
            string lowerBoundString = null;
            string upperBoundString = null;
            string selectionString = null;
            if (id.Equals("Dropdown_None_Upper_Select"))
            {
                lowerBoundString = noneLowerLabel;
                upperBoundString = greenDropdownLabels[greenDropdownIndex];
                selectionString = noneDropdownLabels[index];
            }
            else if (id.Equals("Dropdown_Green_Upper_Select"))
            {
                lowerBoundString = greenLowerLabel;
                upperBoundString = yellowDropdownLabels[yellowDropdownIndex];
                selectionString = greenDropdownLabels[index];
            }
            else if (id.Equals("Dropdown_Yellow_Upper_Select"))
            {
                lowerBoundString = yellowLowerLabel;
                upperBoundString = redUpperLabel;
                selectionString = yellowDropdownLabels[index];
            }

            //get indexes so we can compare
            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            //*IndexOf has linear runtime*
            int lowerBoundIndex = indexNames.IndexOf(lowerBoundString);
            int upperBoundIndex = indexNames.IndexOf(upperBoundString);
            int selectionIndex = indexNames.IndexOf(selectionString);

            //indexOf() returns -1 if not found
            //lowerbound is fine because it is already -1
            if (upperBoundIndex == -1)
            {
                upperBoundIndex = indexNames.Count + 1;
            }

            return EvaluateSelectChange(lowerBoundIndex, selectionIndex, upperBoundIndex);
        }

        private bool EvaluateSelectChange(int lowerBound, int selection, int upperBound)
        {
            bool valid = true;
            if (selection < lowerBound || selection >= upperBound)
            {
                valid = false;
            } //any other error checking??
              //maybe when checking none also check yellow (not just green)

            return valid;
        }

        private void ChangeSelect(string id, int index)
        {
            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            string selection = null;
            int overallIndex = 0;
            if (id.Equals("Dropdown_None_Upper_Select"))
            {
                selection = noneDropdownLabels[index];
                overallIndex = indexNames.IndexOf(selection);
                greenLowerLabel = indexNames[overallIndex + 1];
                noneDropdownIndex = index;
            } else if (id.Equals("Dropdown_Green_Upper_Select"))
            {
                selection = greenDropdownLabels[index];
                overallIndex = indexNames.IndexOf(selection);
                yellowLowerLabel = indexNames[overallIndex + 1];
                greenDropdownIndex = index;
            } else if (id.Equals("Dropdown_Yellow_Upper_Select"))
            {
                selection = yellowDropdownLabels[index];
                overallIndex = indexNames.IndexOf(selection);
                redLowerLabel = indexNames[overallIndex + 1];
                yellowDropdownIndex = index;
            }

            RenderDropdowns();
        }

        internal bool SelectsSet()
        {
            //if none of the selects are enabled
            if (!noneSelectEnabled && !greenSelectEnabled && !yellowSelectEnabled)
                return false;

            //if any of the enabled selects are default value
            if (noneSelectEnabled)
            {
                if (noneDropdownLabels[noneDropdownIndex].Equals(Strings.upperSelectDefault))
                    return false;
            }
            if (greenSelectEnabled)
            {
                if (greenDropdownLabels[greenDropdownIndex].Equals(Strings.upperSelectDefault))
                    return false;
            }
            if (yellowSelectEnabled)
            {
                if (yellowDropdownLabels[yellowDropdownIndex].Equals(Strings.upperSelectDefault))
                    return false;
            }

            //all selects are set
            return true;

        }

        //all info required for highlighting
        internal string GetDropdownNoneLower()
        {
            return noneLowerLabel;
        }

        internal string GetDropdownGreenLower()
        {
            return greenLowerLabel;
        }

        internal string GetDropdownYellowLower()
        {
            return yellowLowerLabel;
        }

        internal string GetDropdownRedLower()
        {
            return redLowerLabel;
        }

        internal string GetDropdownNoneUpper()
        {
            if (noneSelectEnabled)
                return noneDropdownLabels[noneDropdownIndex];
            else
                return null;
        }

        internal string GetDropdownGreenUpper()
        {
            if (greenSelectEnabled)
                return greenDropdownLabels[greenDropdownIndex];
            else
                return null;
        }

        internal string GetDropdownYellowUpper()
        {
            if (yellowSelectEnabled)
                return yellowDropdownLabels[yellowDropdownIndex];
            else
                return null;
        }

        internal string GetDropdownRedUpper()
        {
            if (redUpperLabel.Equals(""))
                return null;
            else
                return redUpperLabel;
        }









        //CONTEXT MENU (right click)
        public void ContextMenuAddWord(Office.IRibbonControl control)
        {
            string selectedText = Globals.ThisAddIn.Application.Selection.Words.First.Text;
            selectedText = Globals.ThisAddIn.HighlightManager.CleanWord(selectedText);
            Globals.ThisAddIn.HighlightManager.AddUnknownWord(this.fileName, selectedText);
        }

        public string ContextMenuLabel(Office.IRibbonControl control)
        {
            string selectedText = Globals.ThisAddIn.Application.Selection.Words.First.Text;
            selectedText = Globals.ThisAddIn.HighlightManager.CleanWord(selectedText);
            selectedText = "'" + selectedText + "'";
            return Strings.sentenceConverter.Reorder("", Strings.addUnknownWordsContextMenuVerb,
                selectedText, Strings.addUnknownWordsContextMenuAdverbPhrase);
        }


        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            ribbon.Invalidate();
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
