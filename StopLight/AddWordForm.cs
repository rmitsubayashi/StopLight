using System;
using System.Collections.Generic;
using System.Linq;
using Forms = System.Windows.Forms;
using System.Drawing;

namespace StopLight
{
    //form to show when the user wants to add an item to the excel file
    internal class AddWordForm : Forms.Form
    {
        private Forms.Button submitButton;
        private Forms.ListBox indexComboBox;
        private Forms.Label text;
        private SelectWordToAddForm form = null;

        private string word = null;
        private string fileName = null;
        

        public AddWordForm()
        {

        }

        public AddWordForm(string wordToAdd, string fileName, SelectWordToAddForm form = null)
        {
            SetWord(wordToAdd);
            SetFileName(fileName);
            InitializeComponents();
            this.form = form;
        }

        private void InitializeComponents()
        {
            if (!WordSet())
                return;
            
            else
            {
                submitButton = new Forms.Button();
                text = new Forms.Label();
                indexComboBox = new Forms.ListBox();

                SetLayout();
                AddEvents();

                Controls.Add(indexComboBox);
                Controls.Add(text);
                Controls.Add(submitButton);
            }
        }

        private void SetLayout()
        {
            int formWidth = 400;
            int formHeight = 350;

            int comboBoxX = 20;
            int comboBoxY = 20;
            int comboBoxWidth = 170;
            int comboBoxHeight = 270;

            int textX = 210;
            int textY = comboBoxY;
            int textWidth = 150;
            int textHeight = 150;

            int buttonX = textX;
            int buttonY = 195;
            int buttonWidth = textWidth;
            int buttonHeight = 90;

            this.Size = new Size(formWidth, formHeight);
            this.Text = Strings.captionAddWord;
            this.FormBorderStyle = Forms.FormBorderStyle.FixedDialog;
            indexComboBox.Location = new Point(comboBoxX, comboBoxY);
            indexComboBox.Size = new Size(comboBoxWidth, comboBoxHeight);
            IList<string> indexNames = Globals.ThisAddIn.HighlightManager.GetIndexNames();
            indexComboBox.Items.AddRange(indexNames.Cast<object>().ToArray()); //needs to be object[]

            text.Size = new Size(textWidth, textHeight);
            text.Location = new Point(textX, textY);
            text.Font = new Font(text.Font.FontFamily, 12);
            text.Text = Strings.sentenceConverter.Reorder("", Strings.messageVerb, "'"+word+"'", Strings.messageAdvP) +
                Environment.NewLine + Strings.messageLine2;

            submitButton.Location = new Point(buttonX, buttonY);
            submitButton.Size = new Size(buttonWidth, buttonHeight);
            submitButton.Font = new Font(submitButton.Font.FontFamily, 20);
            submitButton.Text = Strings.submitButtonText;
            submitButton.FlatStyle = Forms.FlatStyle.Flat;
        }

        private void AddEvents()
        {
            EventHandler handler = new EventHandler(button_Click);
            submitButton.Click += handler;
        }

        private void button_Click(object sender, EventArgs e)
        {
            if (!WordSet())
                return;
            string selection = (string)indexComboBox.SelectedItem;
            if (selection == null)
                return;

            Globals.ThisAddIn.HighlightManager.AddWordToExcel(fileName, word, selection);
            if (DirectedFromForm())
                this.form.RemoveSelectedWord();
            this.Close();
            this.Dispose();
        }
        
        internal void SetWord(string wordToAdd)
        {
            word = wordToAdd;
        }

        private void SetFileName(string fileName)
        {
            this.fileName = fileName;
        }

        private bool WordSet()
        {
            if (word == null)
                return false;
            else
                return true;
        }

        private bool DirectedFromForm()
        {
            if (this.form == null)
                return false;
            else
                return true;
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Forms.Application.Run(new AddWordForm());
        }
    }
}
