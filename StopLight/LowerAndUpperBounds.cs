using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace StopLight
{
    //made a class just to organize code a bit
    internal class LowerAndUpperBounds
    {
        private int noneLowerBound;
        private int greenLowerBound;
        private int yellowLowerBound;
        private int redLowerBound;
        private int noneUpperBound;
        private int greenUpperBound;
        private int yellowUpperBound;
        private int redUpperBound;

        private Ribbon ribbon;
        private IList<String> indexNames;

        public LowerAndUpperBounds(Ribbon ribbon, IList<String> indexNames)
        {
            this.ribbon = ribbon;
            this.indexNames = indexNames;
        }

        internal void FindLowerUpperBounds()
        {
            string tempString; //temp variable
            string selectDefault = Strings.lowerSelectDefault;
            //none
            tempString = ribbon.GetDropdownNoneLower();
            if (tempString.Equals(selectDefault))
                noneLowerBound = -1;
            else //first
                noneLowerBound = 0;

            tempString = ribbon.GetDropdownNoneUpper();
            if (tempString == null)
                noneUpperBound = -1;
            else
                noneUpperBound = indexNames.IndexOf(tempString);

            //green
            tempString = ribbon.GetDropdownGreenLower();
            if (tempString.Equals(selectDefault))
                greenLowerBound = -1;
            else
                greenLowerBound = indexNames.IndexOf(tempString);

            tempString = ribbon.GetDropdownGreenUpper();
            if (tempString == null)
                greenUpperBound = -1;
            else
                greenUpperBound = indexNames.IndexOf(tempString);

            //yellow
            tempString = ribbon.GetDropdownYellowLower();
            if (tempString.Equals(selectDefault))
                yellowLowerBound = -1;
            else
                yellowLowerBound = indexNames.IndexOf(tempString);

            tempString = ribbon.GetDropdownYellowUpper();
            if (tempString == null)
                yellowUpperBound = -1;
            else
                yellowUpperBound = indexNames.IndexOf(tempString);

            tempString = ribbon.GetDropdownRedLower();
            if (tempString.Equals(selectDefault))
                redLowerBound = -1;
            else //last
                redLowerBound = indexNames.IndexOf(tempString);
            tempString = ribbon.GetDropdownRedUpper();
            if (tempString == null)
                redUpperBound = -1;
            else
                redUpperBound = indexNames.Count() - 1;
        }

        internal int NoneLowerBound()
        {
            return noneLowerBound;
        }

        internal int NoneUpperBound()
        {
            return noneUpperBound;
        }

        internal int GreenLowerBound()
        {
            return greenLowerBound;
        }

        internal int GreenUpperBound()
        {
            return greenUpperBound;
        }

        internal int YellowLowerBound()
        {
            return yellowLowerBound;
        }

        internal int YellowUpperBound()
        {
            return yellowUpperBound;
        }

        internal int RedLowerBound()
        {
            return redLowerBound;
        }

        internal int RedUpperBound()
        {
            return redUpperBound;
        }

        internal WdColorIndex GetColorOfIndex(int wordIndex)
        {
            if (wordIndex >= noneLowerBound && wordIndex <= noneUpperBound)
            {
                //wordRange.HighlightColorIndex = none;
                return Colors.none;
            }
            else if (wordIndex >= redLowerBound && wordIndex <= redUpperBound)
            {
                return Colors.red;
            }
            else if (wordIndex >= greenLowerBound && wordIndex <= greenUpperBound)
            {
                return Colors.green;
            }
            else if (wordIndex >= yellowLowerBound && wordIndex <= yellowUpperBound)
            {
                return Colors.yellow;
            }

            return Colors.unknown;
        }

    }
}
