using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KB_XML_Compare
{
    public class DiffTool
    {
        public const string FILE1_PLACEHOLDER = "{1}";
        public const string FILE2_PLACEHOLDER = "{2}";

        protected string m_Caption = "";
        protected string m_Executable = "";
        protected string m_ArgumentTemplate = "";

        public virtual string Caption { get => m_Caption; set => m_Caption = value; }
        public virtual string Executable { get => m_Executable; set => m_Executable = value; }
        public virtual string ArgumentTemplate { get => m_ArgumentTemplate; set => m_ArgumentTemplate = value; }

        public bool VerifyArgPlaceholders()
        {
            return m_ArgumentTemplate.Contains(FILE1_PLACEHOLDER) && m_ArgumentTemplate.Contains(FILE2_PLACEHOLDER);
        }

        public virtual bool RunDiff(string compareFile1, string compareFile2)
        {
            if (!VerifyArgPlaceholders())
            {
                MessageBox.Show("The argument template must include placeholders for the command files being compared.  " +
                                "Please provide a " + FILE1_PLACEHOLDER + " and " + FILE2_PLACEHOLDER + " placeholder in the Diff Tool Args field for command files 1 and 2 respectively.");
                return false;
            }

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = m_Executable.EnsureSurrounded('"');
            compareFile1 = compareFile1.EnsureSurrounded('"');
            compareFile2 = compareFile2.EnsureSurrounded('"');
            startInfo.Arguments = m_ArgumentTemplate.Replace(FILE1_PLACEHOLDER, compareFile1).Replace(FILE2_PLACEHOLDER, compareFile2);

            try
            {
                Process process = Process.Start(startInfo);
                process.WaitForExit();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
            
            return true;
        }
            
    }
}
