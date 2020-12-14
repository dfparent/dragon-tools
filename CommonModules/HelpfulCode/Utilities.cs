using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EquationEditor
{
    class Utilities
    {
        public static string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2}", b);
            return hex.ToString();
        }

        public static byte[] StringToByteArray(String hex)
        {
            int NumberChars = hex.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            return bytes;
        }

        public static string ListToDelimitedString(List<string> fields, string delimiter)
        {
            StringBuilder csv = new StringBuilder();
            foreach (string field in fields)
            {
                string text = field;
                if (text != null)
                {
                    text = text.Replace("\"", "\"\"");
                }
                
                csv.AppendIfNotEmpty(delimiter).Append("\"").Append(text).Append("\"");
            }

            return csv.ToString();
        }

        public static List<string> DelimitedStringToList(string theString, string delimiter)
        {
            List<string> theList = new List<string>();
            bool ignoreDelimiter = false;
            int startIndex = 0;
            int delimiterLength = delimiter.Length;

            int i = 0;
            string newItem = "";

            for (/* Defined above */; i < theString.Length; /* Count inside the loop */)
            {
                if (!ignoreDelimiter && theString.Substring(i, delimiterLength).Equals(delimiter))
                {
                    // End list item
                    // Remove surrounding quotes
                    newItem = theString.Substring(startIndex, i - startIndex);
                    newItem = newItem.RemoveSurrounding("\"");
                    theList.Add(newItem);
                    
                    // Start new list item
                    startIndex = i + delimiterLength;
                    i = startIndex;
                    continue;
                }
                else if (theString.Substring(i, 1).Equals("\""))
                {
                    ignoreDelimiter = !ignoreDelimiter;
                }
                i++;
            }

            // Finish last item
            newItem = theString.Substring(startIndex, i - startIndex);
            newItem = newItem.RemoveSurrounding("\"");
            theList.Add(newItem);
            return theList;
        }
    }
}
