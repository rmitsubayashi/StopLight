using System;
using System.Linq;
using LemmaSharp;

//lemmatizer from http://lemmatise.ijs.si/
//tweaking class to account for minor features I don't want
//cleaner to wrap the class instead of extending??
namespace StopLight
{
    internal class LemmatizerImplementation
    {
        ILemmatizer lemmatizer = null;

        public LemmatizerImplementation()
        {
           lemmatizer = new LemmaSharp.Classes.Lemmatizer(System.IO.File.OpenRead(
               "Lemmas/full7z-mlteast-en.lem")
           );
        }

        internal string Lemmatize(string word)
        {
           string newWord = lemmatizer.Lemmatize(word);
           newWord = RemoveApostrophe(newWord);

           return newWord;
        }

        private String RemoveApostrophe(string word)
        {
            if (word.Equals(""))
                return word;
            //lemmatizing counts words like won't as one word
            //but denotes words like Betty's as Betty' with an apostrophe at the end
            //u2019 is a 'curly quote' for Microsoft Word
            //*u2018 is left quote
            if ((word.Last().Equals('\'') || word.Last().Equals('\u2019')))
            {
                word = word.TrimEnd('\u2019');
                word = word.TrimEnd('\'');
            }
            return word;
        }
    }
}
