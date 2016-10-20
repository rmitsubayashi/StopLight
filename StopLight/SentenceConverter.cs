namespace StopLight
{
    internal class SentenceConverter
    {
        public static readonly string subject_verb_object = "SVO";
        public static readonly string subject_object_verb = "SOV";

        private string order;
        //if the language doesn't have a space
        //space = ""
        string space;

        public SentenceConverter()
        {
            //default (English)
            order = subject_verb_object;
            SetSpacing(true);
        }

        public void SetOrder(string requestedOrder)
        {
            if (!requestedOrder.Equals(subject_object_verb) && 
                !requestedOrder.Equals(subject_verb_object))
            {
                return;
            }

            order = requestedOrder;
        }

        public void SetSpacing(bool spacesRequired)
        {
            if (spacesRequired)
                space = " ";
            else
                space = "";
        }

        public string Reorder(string subject, string verb, string obj, string advP)
        {
            if (order.Equals(subject_object_verb))
                return subject + space + obj + space + advP + space + verb;
            else //default SVO
                return subject + space + verb + space + obj + space + advP;
        }
    }

}
