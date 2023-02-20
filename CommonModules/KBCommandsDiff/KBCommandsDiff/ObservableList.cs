using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KB_XML_Compare
{

    public class ObservableList<T> : BindingList<T> where T : ICloneable
    {
        private T m_AddNewPendingItem = default(T);

        public ObservableList<T> ToDeepList()
        {
            ObservableList<T> newList = new ObservableList<T>();

            foreach (T item in this)
            {
                T newitem = (T)item.Clone();
                newList.Add(newitem);
            }

            if (m_AddNewPendingItem != null)
            {
                newList.m_AddNewPendingItem = (T)m_AddNewPendingItem.Clone();
            }
            else
            {
                newList.m_AddNewPendingItem = default(T);
            }

            // New List will have references to previous event handlers
            newList.ItemAdded = (EventHandler<ObservableListArgs<T>>)ItemAdded.Clone();
            newList.ItemRemoved = (EventHandler<ObservableListArgs<T>>)ItemRemoved.Clone();
            newList.ItemChanged = (EventHandler<ObservableListArgs<T>>)ItemChanged.Clone();
            newList.ListCleared = (EventHandler)ListCleared.Clone();

            return newList;
        }

        public ObservableList() : base()
        {
            
        }

        public ObservableList(IList<T> list) : base(list)
        {
            
        }

        public event EventHandler<ObservableListArgs<T>> ItemAdded;
        public event EventHandler<ObservableListArgs<T>> ItemRemoved;
        public event EventHandler<ObservableListArgs<T>> ItemChanged;
        public event EventHandler ListCleared;

        protected virtual void RaiseEvent(EventHandler<ObservableListArgs<T>> handler, T item)
        {
            handler?.Invoke(this, new ObservableListArgs<T>(item));
        }

        protected virtual void RaiseEvent(EventHandler handler)
        {
            handler?.Invoke(this, EventArgs.Empty);
        }

        public new void Add(T item)
        {
            CheckPendingAddNew();
            base.Add(item);
            RaiseEvent(ItemAdded, item);
        }

        private void CheckPendingAddNew()
        {
            if (m_AddNewPendingItem != null)
            {
                RaiseEvent(ItemAdded, m_AddNewPendingItem);
                m_AddNewPendingItem = default(T);
            }
        }

        public new void AddNew()
        {
            m_AddNewPendingItem = base.AddNew();
        }

        public new void EndNew(int itemIndex)
        {
            base.EndNew(itemIndex);
            if (itemIndex >= 0 && itemIndex < Count)
            {
                RaiseEvent(ItemAdded, Items[itemIndex]);
                m_AddNewPendingItem = default(T);
            }
        }

        public new void CancelNew(int itemIndex)
        {
            base.CancelNew(itemIndex);
            m_AddNewPendingItem = default(T);
        }

        public new void Clear()
        {
            base.Clear();
            RaiseEvent(ListCleared);
            m_AddNewPendingItem = default(T);
        }

        public new bool Remove(T item)
        {
            CheckPendingAddNew();

            if (base.Remove(item))
            {
                RaiseEvent(ItemRemoved, item);
                return true;
            }
            else
            {
                return false;
            }
        }

        public new void RemoveAt(int index)
        {
            CheckPendingAddNew();

            T item = default(T);
            if (index >= 0 && index < Count)
            {
                item = Items[index];
                base.RemoveAt(index);
                RaiseEvent(ItemRemoved, item);
            }
            else
            {
                // Will throw exception
                base.RemoveAt(index);
            }
        }

        public new void ResetItem(int index)
        {
            if (index >= 0 && index < Count)
            {
                T item = Items[index];
                base.ResetItem(index);
                RaiseEvent(ItemChanged, item);
            }
            else
            {
                base.ResetItem(index);
            }
        }

        public new void Insert(int index, T item)
        {
            CheckPendingAddNew();

            if (index >= 0 && index <= Count)
            {
                base.Insert(index, item);
                RaiseEvent(ItemAdded, item);
            }
            else
            {
                // Will throw exception
                base.Insert(index, item);
            }
        }

        


    }

    public class ObservableListArgs<T> : EventArgs
    {
        public ObservableListArgs(T item)
        {
            Item = item;
        }

        public T Item { get; }
    }

}
