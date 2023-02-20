using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KB_XML_Compare
{
    class DiffToolBeyondCompare : DiffTool
    {
        public DiffToolBeyondCompare()
        {
            m_Caption = "Beyond Compare";
            m_Executable = @"C:\Program Files\Beyond Compare 4\BCompare.exe";
            m_ArgumentTemplate = "{1} {2}";
        }

    }
}
