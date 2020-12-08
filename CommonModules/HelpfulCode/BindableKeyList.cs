using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EquationEditor
{
    public class GenericIListToIList<T> : IList
    {
        private IList<T> m_genericIList = null;

        public GenericIListToIList(IList<T> genericIList)
        {
            m_genericIList = genericIList;
        }

        private GenericIListToIList()
        {

        }

        public object this[int index]
        {
            get => m_genericIList.ElementAt(index);
            set => m_genericIList.Insert(index, (T)value);
        }

        public bool IsReadOnly => m_genericIList.IsReadOnly;

        public bool IsFixedSize => false;

        public int Count => m_genericIList.Count;

        public object SyncRoot => throw new NotImplementedException();

        public bool IsSynchronized => throw new NotImplementedException();

        public int Add(object value)
        {
            m_genericIList.Add((T) value);
            return m_genericIList.IndexOf((T)value);
        }

        public void Clear()
        {
            m_genericIList.Clear();
        }

        public bool Contains(object value)
        {
            return m_genericIList.Contains((T)value);
        }

        public void CopyTo(Array array, int index)
        {
            throw new NotImplementedException();
        }

        public IEnumerator GetEnumerator()
        {
            return m_genericIList.GetEnumerator();
        }

        public int IndexOf(object value)
        {
            return m_genericIList.IndexOf((T)value);
        }

        public void Insert(int index, object value)
        {
            m_genericIList.Insert(index, (T) value);
        }

        public void Remove(object value)
        {
            m_genericIList.Remove((T) value);
        }

        public void RemoveAt(int index)
        {
            m_genericIList.RemoveAt(index);
        }
    }
}
