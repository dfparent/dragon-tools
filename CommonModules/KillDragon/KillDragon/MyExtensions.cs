using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Reflection;
using System.Collections.ObjectModel;
//using System.Windows.Forms;

namespace KillDragon
{
    public static class MyExtensions
    {
        /*************************/
        /* Enum stuff */

        public static string GetDescription(this Enum value)
        {
            FieldInfo field = value.GetType().GetField(value.ToString());
            DescriptionAttribute attribute = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;
            return attribute == null ? value.ToString() : attribute.Description;
        }

        public static Enum GetEnumValueByDescription(Type theType, string desc)
        {
            if (!theType.IsEnum)
            {
                throw new Exception("The provided type is not an Enum: " + theType.Name);
            }

            foreach (Enum aValue in Enum.GetValues(theType))
            {
                if (aValue.GetDescription().Equals(desc))
                {
                    return aValue;
                }
            }

            throw new Exception("Invalid value for " + theType.Name + ": " + desc);
        }
        
        /*************************/
        /* String stuff */

        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source != null && toCheck != null && source.IndexOf(toCheck, comp) >= 0;
        }


        public static StringBuilder AppendIfNotEmpty(this StringBuilder builder, string appendString)
        {
            if (builder.Length != 0)
            {
                builder.Append(appendString);
            }
            return builder;
        }

        public static string RemoveSurrounding(this string theString, string surroundingString)
        {
            if (theString.StartsWith(surroundingString) && theString.EndsWith(surroundingString))
            {
                theString = theString.Substring(surroundingString.Length, theString.Length - (surroundingString.Length * 2));
            }
            return theString;
        }

        /*************************/
        /* BindingList stuff */

        public static BindingList<T> CastBindingList<F, T>(this BindingList<F> fromList)
        {

            return new BindingList<T>(fromList.Cast<T>().ToList());

        }

        public static BindingList<T> ToBindingList<T>(this IEnumerable<T> fromEnum)
        {
            BindingList<T> newList = new BindingList<T>();
            foreach (T item in fromEnum)
            {
                newList.Add(item);
            }
            return newList;
        }

        public static BindingList<T> ToDeepBindingList<T>(this IEnumerable<T> fromEnum) where T : ICloneable
        {
            BindingList<T> newList = new BindingList<T>();
            
            foreach (T item in fromEnum)
            {
                T newitem = (T)item.Clone();
                newList.Add(newitem);
            }

            return newList;
        }

        public static void AddRange<T>(this BindingList<T> theBindingList, IEnumerable<T> addCollection )
        {
            foreach (T item in addCollection)
            {
                theBindingList.Add(item);
            }
        }

        public static void Resize<T>(this List<T> theList, int numberOfElements, T placeHolder)
        {
            if (theList.Count == numberOfElements)
            {
                return;
            }
            else if (theList.Count > numberOfElements)
            {
                // Remove last items
                for (int i = theList.Count - 1; i >= numberOfElements; i--)
                {
                    theList.RemoveAt(i);
                }
            }
            else
            {
                // Add items at end
                for (int i = theList.Count; i < numberOfElements; i++)
                {
                    theList.Add(placeHolder);
                }
            }
        }

        /*************************/
        /* Text Box stuff */

        // Returns location of found text or -1 if not found
        /*public static int Find(this TextBox theBox, string searchText, int start = 0)
        {
            int index = theBox.Text.IndexOf(searchText, start);
            if (index >= 0)
            {
                theBox.SelectionStart = index;
                theBox.SelectionLength = searchText.Length;
            }

            return index;
        }*/
    }
}
