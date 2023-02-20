using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Reflection;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace KB_XML_Compare
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

        public enum StringCase
        {
            AllUpperCase,
            AllLowerCase,
            InitialCap,
            MixedCase
        }

        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source != null && toCheck != null && source.IndexOf(toCheck, comp) >= 0;
        }

        public static StringCase GetCase(this string theString)
        {
            bool hasUpper = false;
            bool hasLower = false;
            bool hasInitialUpperOnly = false;

            char[] chars = theString.ToCharArray();

            for (int i = 0; i < chars.Length; i++)
            {
                if (char.IsUpper(chars[i]))
                {
                    hasUpper = true;
                    if (i == 0)
                    {
                        hasInitialUpperOnly = true;
                    }
                    else
                    {
                        hasInitialUpperOnly = false;
                    }
                }
                else if (char.IsLower(chars[i]))
                {
                    hasLower = true;
                }
            }

            if (hasUpper && !hasLower)
            {
                return StringCase.AllUpperCase;
            }
            else if (hasLower && !hasUpper)
            {
                return StringCase.AllLowerCase;
            }
            else if (hasInitialUpperOnly)
            {
                return StringCase.InitialCap;
            }
            else
            {
                return StringCase.MixedCase;
            }
        }

        public static string ToInitialCap(this string theString)
        {
            // Check inputs.
            if (theString == null)
            {
                // Same as original .NET C# string.Replace behavior.
                throw new ArgumentNullException(nameof(theString));
            }
            if (theString.Length == 0)
            {
                // Same as original .NET C# string.Replace behavior.
                return theString;
            }

            if (theString.Length == 1)
            {
                return char.ToUpper(theString[0]).ToString();
            }
            else
            {
                return char.ToUpper(theString[0]) + theString.Substring(1).ToLower();
            }
            
            
        }

        public static bool IsNullOrEmpty(this string theString)
        {
            if (theString == null || theString.Length == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Ensures that the string starts with and ends with the indicated surrounding character and returns the result.
        /// If the original string is already surrounded by the indicated character, then the string is returned unchanged.
        /// If the surrounding character appears in the string it is escaped by doubling it up.
        /// </summary>
        /// <param name="theString"></param>
        /// <param name="surroundingChar"></param>
        /// <returns></returns>
        public static string EnsureSurrounded(this string theString, char surroundingChar)
        {
            if (theString == null || theString.Length == 0)
            {
                return theString;
            }

            string surroundString = surroundingChar.ToString();
            StringBuilder newString = new StringBuilder(theString);

            newString = newString.Replace(surroundString, surroundString + surroundString);
            if (!theString.StartsWith(surroundString))
            {
                newString.Insert(0, surroundString);
            }

            if (!theString.EndsWith(surroundString))
            {
                newString.Append(surroundString);
            }

            return newString.ToString();
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

        // From https://stackoverflow.com/questions/6275980/string-replace-ignoring-case
        public static string Replace(this string str, string oldValue, string newValue, StringComparison comparisonType, bool matchCaseOnReplace = false)
        {
            // Check inputs.
            if (str == null)
            {
                // Same as original .NET C# string.Replace behavior.
                throw new ArgumentNullException(nameof(str));
            }
            if (str.Length == 0)
            {
                // Same as original .NET C# string.Replace behavior.
                return str;
            }
            if (oldValue == null)
            {
                // Same as original .NET C# string.Replace behavior.
                throw new ArgumentNullException(nameof(oldValue));
            }
            if (oldValue.Length == 0)
            {
                // Same as original .NET C# string.Replace behavior.
                throw new ArgumentException("String cannot be of zero length.");
            }

            // Prepare string builder for storing the processed string.
            // Note: StringBuilder has a better performance than String by 30-40%.
            StringBuilder resultStringBuilder = new StringBuilder(str.Length);

            // Analyze the replacement: replace or remove.
            bool isReplacementNullOrEmpty = string.IsNullOrEmpty(newValue);

            // Replace all values.
            const int valueNotFound = -1;
            int foundAt;
            int startSearchFromIndex = 0;
            while ((foundAt = str.IndexOf(oldValue, startSearchFromIndex, comparisonType)) != valueNotFound)
            {
                // Append all characters until the found replacement.
                int charsUntilReplacment = foundAt - startSearchFromIndex;
                bool isNothingToAppend = charsUntilReplacment == 0;
                if (!isNothingToAppend)
                {
                    resultStringBuilder.Append(str, startSearchFromIndex, charsUntilReplacment);
                }

                // Process the replacement.
                if (!isReplacementNullOrEmpty)
                {
                    if (matchCaseOnReplace)
                    {
                        // Make the replacement text match the case of the text being replaced
                        string foundText = str.Substring(foundAt, oldValue.Length);
                        StringCase stringCase = foundText.GetCase();
                        string matchedCaseNewValue = null;
                        switch (stringCase)
                        {
                            case StringCase.AllUpperCase:
                                matchedCaseNewValue = newValue.ToUpper();
                                break;
                            case StringCase.AllLowerCase:
                                matchedCaseNewValue = newValue.ToLower();
                                break;
                            case StringCase.InitialCap:
                                matchedCaseNewValue = newValue.ToInitialCap();
                                break;
                            case StringCase.MixedCase:
                                matchedCaseNewValue = newValue;
                                break;
                            default:
                                break;
                        }
                        resultStringBuilder.Append(matchedCaseNewValue);
                    }
                    else
                    {
                        resultStringBuilder.Append(newValue);
                    }
                    
                }

                // Prepare start index for the next search.
                // This needed to prevent infinite loop, otherwise method always start search 
                // from the start of the string. For example: if an oldValue == "EXAMPLE", newValue == "example"
                // and comparisonType == "any ignore case" will conquer to replacing:
                // "EXAMPLE" to "example" to "example" to "example" … infinite loop.
                startSearchFromIndex = foundAt + oldValue.Length;
                if (startSearchFromIndex == str.Length)
                {
                    // It is end of the input string: no more space for the next search.
                    // The input string ends with a value that has already been replaced. 
                    // Therefore, the string builder with the result is complete and no further action is required.
                    return resultStringBuilder.ToString();
                }
            }


            // Append the last part to the result.
            int charsUntilStringEnd = str.Length - startSearchFromIndex;
            resultStringBuilder.Append(str, startSearchFromIndex, charsUntilStringEnd);


            return resultStringBuilder.ToString();

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

        public static void AddRange<T>(this ObservableList<T> theList, IEnumerable<T> addCollection) where T : ICloneable
        {
            foreach (T item in addCollection)
            {
                theList.Add(item);
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

        public static void Grow<T>(this Collection<T> theCollection, int newSize, T placeHolder)
        {
            if (theCollection.Count >= newSize)
            {
                return;
            }

            // Add items at end
            for (int i = theCollection.Count; i < newSize; i++)
            {
                theCollection.Add(placeHolder);
            }
        }

        /*************************/
        /* List stuff */

        public static List<T> ToDeepList<T>(this IEnumerable<T> fromEnum) where T : ICloneable
        {
            List<T> newList = new List<T>();

            foreach (T item in fromEnum)
            {
                T newItem = (T)item.Clone();
                newList.Add(newItem);
            }
            return newList;
        }

        /*************************/
        /* HashSet stuff */

        public static HashSet<T> ToDeepHashSet<T>(this IEnumerable<T> fromEnum) where T : ICloneable
        {
            HashSet<T> newHashSet = new HashSet<T>();

            foreach (T item in fromEnum)
            {
                T newItem = (T)item.Clone();
                newHashSet.Add(newItem);
            }
            return newHashSet;
        }

        
        /// <summary>
        /// Adds Multiple elements to a hash set.  A HashSet can only contain unique elements,
        /// so this method will only add elements that are not already in the HashSet.
        /// No error is thrown if an element is already in the set.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="theHashSet"></param>
        /// <param name="fromEnum"></param>
        public static void AddRange<T>(this HashSet<T> theHashSet, IEnumerable<T> fromEnum)
        {
            foreach (T item in fromEnum)
            {
                theHashSet.Add(item);
            }
        }

        /*************************/
        /* Array stuff */

        public static T[] ToDeepArray<T>(this T[] fromArray) where T : ICloneable
        {
            T[] newArray = new T[fromArray.Count()];

            for (int i = 0; i < fromArray.Count(); i++)
            {
                newArray[i] = (T)(fromArray[i].Clone());
            }

            return newArray;
        }

        /*************************/
        /* Text Box stuff */

        // Returns location of found text or -1 if not found
        public static int Find(this TextBox theBox, string searchText, int start = 0)
        {
            int index = theBox.Text.IndexOf(searchText, start);
            if (index >= 0)
            {
                theBox.SelectionStart = index;
                theBox.SelectionLength = searchText.Length;
            }

            return index;
        }

        // Code from stack overflow: https://stackoverflow.com/questions/4683663/how-to-remove-annoying-beep-with-richtextbox

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private extern static IntPtr SendMessage(HandleRef hWnd, Int32 msg, IntPtr wParam, ref ITextServices lParam);

        #region ITextServices 
        // From TextServ.h
        [ComImport(), Guid("8d33f740-cf58-11ce-a89d-00aa006cadc5"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ITextServices
        {
            //see: Slots in the V-table
            //     https://docs.microsoft.com/en-us/dotnet/framework/unmanaged-api/metadata/imetadataemit-definemethod-method#slots-in-the-v-table
            void _VtblGap1_16();
            Int32 OnTxPropertyBitsChange(TXTBIT dwMask, Int32 dwBits);
        }

        private enum TXTBIT : uint
        {
            /// <summary>If TRUE, beeping is enabled.</summary>
            ALLOWBEEP = 0x800
        }
        #endregion

        public static void EnableBeep(this RichTextBox rtb)
        {
            SetBeepInternal(rtb, true);
        }

        public static void DisableBeep(this RichTextBox rtb)
        {
            SetBeepInternal(rtb, false);
        }

        private static void SetBeepInternal(RichTextBox rtb, bool beepOn)
        {
            const Int32 WM_USER = 0x400;
            const Int32 EM_GETOLEINTERFACE = WM_USER + 60;
            const Int32 COMFalse = 0;
            const Int32 COMTrue = ~COMFalse; // -1

            ITextServices textServices = null;
            // retrieve the rtb's OLEINTERFACE using the Interop Marshaller to cast it as an ITextServices
            // The control calls the AddRef method for the object before returning, so the calling application must call the Release method when it is done with the object.
            SendMessage(new HandleRef(rtb, rtb.Handle), EM_GETOLEINTERFACE, IntPtr.Zero, ref textServices);
            textServices.OnTxPropertyBitsChange(TXTBIT.ALLOWBEEP, beepOn ? COMTrue : COMFalse);
            Marshal.ReleaseComObject(textServices);
        }

        /**********************************/
        /*  DataGridView Stuff  */

        /// <summary>
        /// Turn on to improve scrolling performance in grids with a lot of rows
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="setting"></param>
        public static void SetDoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }

        /// <summary>
        /// Returns the state of the DoubleBuffered property.
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="setting"></param>
        public static bool GetDoubleBuffered(this DataGridView dgv)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            return (bool)pi.GetValue(dgv);
        }

    }
}
