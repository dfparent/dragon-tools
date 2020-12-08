using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EquationEditor
{
    // Adds the IList interface to the SortedList class So that it can be bound to the datagridview
    public class BindableSortedList<TKey, TValue> : SortedList<TKey, TValue>
    {
        GenericIListToIList<TKey> m_genericKeyList = null;
        GenericIListToIList<TValue> m_genericValueList = null;

        public BindableSortedList() : base()
        {
            m_genericKeyList = new GenericIListToIList<TKey>(Keys);
            m_genericValueList = new GenericIListToIList<TValue>(Values);
        }

        public GenericIListToIList<TKey> GetBindableKeyIList()
        {
            return m_genericKeyList;
        }

        public GenericIListToIList<TValue> GetBindableValueIList()
        {
            return m_genericValueList;
        }

    }
}
