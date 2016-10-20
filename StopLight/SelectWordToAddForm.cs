using System;
using System.Collections.Generic;
using System.Linq;
using Forms = System.Windows.Forms;
using System.Drawing;

namespace StopLight
{
    internal class SelectWordToAddForm : Forms.Form
    {
        private Forms.Label text;
        private Forms.Button submitButton;
        private Forms.ListBox unknownWordsComboBox;


        private string fileName = null;

        public SelectWordToAddForm()
        {

        }

        public SelectWordToAddForm(string fileName)
        {
            SetFileName(fileName);
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            text = new Forms.Label();
            unknownWordsComboBox = new Forms.ListBox();
            submitButton = new Forms.Button();

            SetLayout();
            AddEvents();
            
            this.Controls.Add(unknownWordsComboBox);
            this.Controls.Add(text);
            this.Controls.Add(submitButton);
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
            this.Text = Strings.captionSelectWord;
            this.FormBorderStyle = Forms.FormBorderStyle.FixedDialog;
            unknownWordsComboBox.Location = new Point(comboBoxX, comboBoxY);
            unknownWordsComboBox.Size = new Size(comboBoxWidth, comboBoxHeight);
            IList<string> unknownWords = Globals.ThisAddIn.HighlightManager.GetUnknownWords();
            unknownWordsComboBox.Items.AddRange(unknownWords.Cast<object>().ToArray()); //needs to be object[]

            text.Size = new Size(textWidth, textHeight);
            text.Location = new Point(textX, textY);
            text.Text = Strings.message;
            text.Font = new Font(text.Font.FontFamily, 12);

            submitButton.Location = new Point(buttonX, buttonY);
            submitButton.Size = new Size(buttonWidth, buttonHeight);
            submitButton.Text = Strings.selectButtonText;
            submitButton.Font = new Font(submitButton.Font.FontFamily, 20);
            submitButton.FlatStyle = Forms.FlatStyle.Flat;


            //Forms.MessageBox.Show("Button (" + buttonX + "," + buttonY + ") WH=" + buttonWidth + " x " +
            //    buttonHeight + " Combo Box (" + comboBoxX + "," + comboBoxY + ") WH=" + comboBoxWidth + " x " + comboBoxHeight);
            
        }

        private void AddEvents()
        {
            EventHandler handler = new EventHandler(button_Click);
            submitButton.Click += handler;
        }

        private void button_Click(object sender, EventArgs e)
        {
            string selection = (string)unknownWordsComboBox.SelectedItem;
            if (selection == null)
                return;

            Globals.ThisAddIn.HighlightManager.AddUnknownWord(fileName, selection, this);
            //after removing, the form will call RemoveSelectedWord()
        }

        internal void RemoveSelectedWord()
        {
            //since the selected item has to be the item we removed
            //just remove the selected item
            unknownWordsComboBox.Items.Remove(unknownWordsComboBox.SelectedItem);

        }

        private void SetFileName(string fileName)
        {
            this.fileName = fileName;
        }
        

        [STAThread]
        public static void Main(string[] args)
        {
            Forms.Application.Run(new SelectWordToAddForm());
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // SelectWordToAddForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "SelectWordToAddForm";
            this.Load += new System.EventHandler(this.SelectWordToAddForm_Load);
            this.ResumeLayout(false);

        }

        private void SelectWordToAddForm_Load(object sender, EventArgs e)
        {

        }
    }
}
