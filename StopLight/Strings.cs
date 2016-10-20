using System;

namespace StopLight
{
    //provides every string in the UI
    //plus language compatibility
    internal class Strings
    {
        public static SentenceConverter sentenceConverter = new SentenceConverter();

        //General
        public static String error = "An error has occured";
        public static String errorCaption = "Error!";

        //
        // Main Class
        //
        public static String alreadyHighlightedMessage = "Highlighted text will temporarily be overwritten. Do you want to continue?";
        public static String alreadyHighlightedCaption = "Warning: Existing Highlighted Text!!";
        //default select option
        public static String selectNothing = "--Select--";
        //a message box that appears when the user tries highlighting when the selections aren't set
        public static String rangeNotSetMessage = "The ranges are not set yet. Please set the proper ranges.";
        public static String rangeNotSetCaption = "Error: Range Not Set!!";
        //when the user tries an action without selecting a file
        public static String fileNotSelectedMessage = "Please select a file first!";
        public static String fileNotSelectedCaption = "File Not Selected";

        //
        //Ribbon
        //
        //a message box that pops up when the user enters an invalid file
        public static String cannotReadFileMessage = "The file could not be read. Please try again.";
        public static String cannotReadFileCaption = "Error: Cannot Read File!!";
        public static String actionsGroup = "Action";
        public static String highlightButton = "Highlight";
        public static String removeButton = "Remove";

        public static String selectFileGroup = "Data File";
        public static String selectFileButton = "Select Excel File";

        //fileNameLabel will also have the actual file name!
        //because ribbon won't allow me to put a label right next to it
        //make sure the colon is an alphanumeric one
        //so we can add on to it
        public static String fileNameLabel = "File Name: ";// + file name

        //default colon is "  :  " preceded and followed by spaces
        public static String dropdownNoneLabel = "None  ";
        public static String dropdownGreenLabel = "Green  ";
        public static String dropdownYellowLabel = "Yellow  ";
        public static String dropdownRedLabel = "Red  ";
        public static String rangeGroup = "Range";

        public static String addUnknownWordsGroup = "Unknown Words";
        public static String addUnknownWordsButton = "Add to Excel File";
        public static String addUnknownWordsContextMenuVerb = "Add";
        public static String addUnknownWordsContextMenuAdverbPhrase = "to Excel File";

        //default select option
        public static String upperSelectDefault = "--Select--";
        public static readonly String lowerSelectDefault = "  ?  ";

        //
        //Add Word Form
        //
        public static string captionAddWord = "Select Column";
        public static string messageVerb = "Adding";
        public static string messageAdvP = "to Excel file.";
        public static string messageLine2 = "Please select a column.";
        public static string submitButtonText = "Submit";

        public static string duplicateWordCaption = "Error";
        public static string duplicateWordText = "The word already exists in the Excel file.";

        //
        //Select Word to Add Form
        //
        public static string captionSelectWord = "Select Word";
        public static string message = "Please select the word you would like to add to the Excel file.";
        public static string selectButtonText = "Select";

        internal static void Japanese()
        {
            sentenceConverter.SetOrder(SentenceConverter.subject_object_verb);
            sentenceConverter.SetSpacing(true);
            //General
            error = "エラーが起こりました";
            errorCaption = "エラー";

            //
            //Main class
            //
            alreadyHighlightedMessage = "一時的に既存の蛍光ペンは上書きされます。続行しますか？";
            alreadyHighlightedCaption = "注意：既存の蛍光ペンが存在します！！";

            selectNothing = "--選択--";

            rangeNotSetMessage = "範囲を設定してからもう一度お試しください";
            rangeNotSetCaption = "エラー：範囲が設定されていません";

            fileNotSelectedMessage = "まずファイルを選択してください。";
            fileNotSelectedCaption = "ファイルが選択されていません";

            //
            //Ribbon
            //
            //this is a brand name so no need to make Japanese??
            //this.StopLightTab.Label = "ストップライト";

            actionsGroup = "実行";
            highlightButton = "蛍光ペン";
            removeButton = "消去";

            selectFileGroup = "データファイル";
            selectFileButton = "Excelファイルを選択";
            //make sure the colon is an alphanumeric one
            //so we can add on to it
            fileNameLabel = "ファイル名: ";
            
            dropdownNoneLabel = "なし　";
            dropdownGreenLabel = "みどり";
            dropdownYellowLabel = "きいろ";
            dropdownRedLabel = "あか　";
            rangeGroup = "範囲";

            addUnknownWordsGroup = "不明な単語";
            addUnknownWordsButton = "Excelファイルに追加";
            addUnknownWordsContextMenuVerb = "に足す";
            addUnknownWordsContextMenuAdverbPhrase = "をエクセルファイル";

            upperSelectDefault = "--選択--";

            cannotReadFileMessage = "ファイルの読み込みに失敗しました。もう一度お試しください。";
            cannotReadFileCaption = "エラー：ファイルの読み込み失敗";

            //
            //Add Word Form
            //
            captionAddWord = "単語の選択";
            selectButtonText = "選択";
            message = "Excelファイルに追加する単語を選択してください";

            duplicateWordCaption = "エラー";
            duplicateWordText = "この単語はすでにExcelファイルに存在しています";

            //
            //Select Word to Add Form
            //
            captionSelectWord = "列の選択";
            submitButtonText = "送信";
            messageAdvP = "をExcelファイル";
            messageVerb = "に追加します。";
            messageLine2 = "列を選択してください";
            
        }
        
    }
}
