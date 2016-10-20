using Word = Microsoft.Office.Interop.Word;

namespace StopLight
{
    public static class Colors
    {
        public const Word.WdColorIndex red = Word.WdColorIndex.wdRed;
        public const Word.WdColorIndex yellow = Word.WdColorIndex.wdYellow;
        public const Word.WdColorIndex green = Word.WdColorIndex.wdBrightGreen;
        public const Word.WdColorIndex none = Word.WdColorIndex.wdNoHighlight;
        public const Word.WdColorIndex unknown = Word.WdColorIndex.wdGray25;
    }
}
