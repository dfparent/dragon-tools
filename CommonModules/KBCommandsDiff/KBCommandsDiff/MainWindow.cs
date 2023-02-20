using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace KB_XML_Compare
{
    public partial class mainWindow : Form
    {
        public const string DEFAULT_KB_COMMANDS_FILE = "MyKBCommands.xml";
        public const string DIFF_TOOL_CUSTOM_NAME = "Custom";

        private Dictionary<string, DiffTool> m_DiffTools = new Dictionary<string, DiffTool>();
        private XElement m_xmlRootElement1 = null;
        private XElement m_xmlRootElement2 = null;

        public mainWindow()
        {
            InitializeComponent();

            // Populate diff tools

            // Beyond Compare
            DiffTool diffTool = new DiffToolBeyondCompare();
            m_DiffTools.Add(diffTool.Caption, diffTool);
            cboDiffTool.Items.Add(diffTool.Caption);

            // Win Merge
            diffTool = new DiffToolWinMerge();
            m_DiffTools.Add(diffTool.Caption, diffTool);
            cboDiffTool.Items.Add(diffTool.Caption);

            // Custom
            diffTool = new DiffTool();
            diffTool.Caption = DIFF_TOOL_CUSTOM_NAME;
            diffTool.Executable = Properties.Settings.Default.CustomDiffToolExe;
            diffTool.ArgumentTemplate = Properties.Settings.Default.CustomDiffToolArgs;
            m_DiffTools.Add(diffTool.Caption, diffTool);
            cboDiffTool.Items.Add(diffTool.Caption);

            LoadSettings();

            UpdateCompareButtonEnableState();
            UpdateDiffToolFieldsEnableState();
        }

        private void LoadSettings()
        {
            txtCommandFile1.Text = Properties.Settings.Default.CommandFile1;
            txtCommandFile2.Text = Properties.Settings.Default.CommandFile2;
            if (cboDiffTool.Items.Contains(Properties.Settings.Default.DiffToolType))
            {
                cboDiffTool.Text = Properties.Settings.Default.DiffToolType;
                if (cboDiffTool.Text.Equals(DIFF_TOOL_CUSTOM_NAME))
                {
                    txtDiffToolArgs.Text = Properties.Settings.Default.CustomDiffToolArgs;
                    txtDiffToolExe.Text = Properties.Settings.Default.CustomDiffToolExe;
                }
            }
            else
            {
                cboDiffTool.Text = "";
            }

            chkDeleteTempFiles.Checked = Properties.Settings.Default.DeleteTempFiles;
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.CommandFile1 = txtCommandFile1.Text;
            Properties.Settings.Default.CommandFile2 = txtCommandFile2.Text;
            Properties.Settings.Default.DiffToolType = cboDiffTool.Text;
            Properties.Settings.Default.CustomDiffToolArgs = m_DiffTools[DIFF_TOOL_CUSTOM_NAME].ArgumentTemplate;
            Properties.Settings.Default.CustomDiffToolExe = m_DiffTools[DIFF_TOOL_CUSTOM_NAME].Executable;
            Properties.Settings.Default.DeleteTempFiles = chkDeleteTempFiles.Checked;
            Properties.Settings.Default.Save();
        }

        private void btnCommandFile1_Click(object sender, EventArgs e)
        {
            HandleXmlFileSelect(txtCommandFile1);
        }

        private void btnCommandFile2_Click(object sender, EventArgs e)
        {
            HandleXmlFileSelect(txtCommandFile2);
        }

        public void HandleXmlFileSelect(TextBox theTextBox)
        {
            openFileDialog.DefaultExt = ".xml";
            openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
            openFileDialog.FileName = DEFAULT_KB_COMMANDS_FILE;

            // FileInfo fileInfo = new FileInfo(openFileDialog.FileName);
            // openFileDialog.InitialDirectory = fileInfo.DirectoryName;

            if (openFileDialog.ShowDialog(this) == DialogResult.Cancel)
            {
                return;
            }

            if (openFileDialog.FileName.Equals(""))
            {
                return;
            }

            theTextBox.Text = openFileDialog.FileName;
        }

        public void HandleExeFileSelect()
        {
            openFileDialog.DefaultExt = ".exe";
            openFileDialog.Filter = "EXE Files (*.exe)|*.exe|All Files (*.*)|*.*";

            // FileInfo fileInfo = new FileInfo(openFileDialog.FileName);
            // openFileDialog.InitialDirectory = fileInfo.DirectoryName;

            if (openFileDialog.ShowDialog(this) == DialogResult.Cancel)
            {
                return;
            }

            if (openFileDialog.FileName.Equals(""))
            {
                return;
            }

            txtDiffToolExe.Text = openFileDialog.FileName;
        }

        private void btnCompare_Click(object sender, EventArgs e)
        {
            DoCompare();
        }

        private void DoCompare()
        {
            Cursor = Cursors.WaitCursor;

            try
            {
                // Load XML
                m_xmlRootElement1 = ReadXmlFile(txtCommandFile1.Text);
                m_xmlRootElement2 = ReadXmlFile(txtCommandFile2.Text);

                // Normalize XML
                XElement newRootElement1 = NormalizeXml(m_xmlRootElement1);
                XElement newRootElement2 = NormalizeXml(m_xmlRootElement2);

                if (!WriteTempXmlFile(newRootElement1, txtCommandFile1.Text, out string normalizedTempFile1)) return;
                if (!WriteTempXmlFile(newRootElement2, txtCommandFile2.Text, out string normalizedTempFile2)) return;

                // Diff XML
                m_DiffTools[cboDiffTool.Text].RunDiff(normalizedTempFile1, normalizedTempFile2);

                if (chkDeleteTempFiles.Checked)
                {
                    File.Delete(normalizedTempFile1);
                    File.Delete(normalizedTempFile2);
                }

            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private XElement ReadXmlFile(string filePath)
        {
            try
            {
                return XElement.Load(filePath);
            }
            catch (Exception e)
            {
                MessageBox.Show(this, e.Message);
                return null;
            }
        }

        private bool WriteTempXmlFile(XElement xmlRootElement, string originalFilePath, out string newTempFilePath)
        {
            StringBuilder tempFilePath = new StringBuilder();
            tempFilePath.Append(Path.GetDirectoryName(originalFilePath)).Append(Path.DirectorySeparatorChar)
                .Append(Path.GetFileNameWithoutExtension(originalFilePath))
                .Append("_").Append(DateTime.Now.Ticks)
                .Append(Path.GetExtension(originalFilePath));
            try
            {
                xmlRootElement.Save(tempFilePath.ToString());
                newTempFilePath = tempFilePath.ToString();
                return true;
            }
            catch(Exception e)
            {
                MessageBox.Show(this, "Unable to save XML file: " + tempFilePath.ToString());
                newTempFilePath = null;
                return false;
            }
            
        }
        // Sorts elements and attributes so that a file compare can be reasonably performed
        // File format is:
        /**
         * <KnowBrainerCommands>
         *      <Commands scope="global">  <!-- Sort on scope -->
         *          <Command name="">      <!-- Sort on name -->
         *          </Command>
         *      </Commands>
         *      <Lists>
         *          <list name="">         <!-- Sort on name -->
         *          </list>
         *      </Lists>
         *      <Commands scope="application" module="">  <!-- Sort on scope, module -->
         *          <Command name="">       <!-- Sort on name -->
         *          </Command>
         *      </Commands>
         *      <Commands scope="window" module="" windowTitle="">      <!-- Sort on scope, module, windowTitle -->
         *      </Commands>
         * </KnowBrainerCommands>
         **/
        private XElement NormalizeXml(XElement xmlRootElement)
        {
            XElement newRootElement = new XElement("KnowBrainerCommands");

            // Add and sort global commands
            IEnumerable<XElement> commandsElements = xmlRootElement.Descendants("Commands").Where(item => item.Attribute("scope").Value == "global");
            foreach(XElement elem in commandsElements)
            {
                elem.ReplaceNodes(elem.Descendants("Command").OrderBy(item => item.Attribute("name").Value));
                newRootElement.Add(elem);
            }

            // Add and sort application commands
            commandsElements = xmlRootElement.Descendants("Commands")
                .Where(item => item.Attribute("scope").Value == "application")
                .OrderBy(item => item.Attribute("module").Value);
            foreach(XElement elem in commandsElements)
            {
                elem.ReplaceNodes(elem.Descendants("Command").OrderBy(item => item.Attribute("name").Value));
                newRootElement.Add(elem);
            }

            // Add and sort window commands
            commandsElements = xmlRootElement.Descendants("Commands")
                .Where(item => item.Attribute("scope").Value == "window")
                .OrderBy(item => item.Attribute("module").Value).ThenBy(item => item.Attribute("windowTitle").Value);
            foreach (XElement elem in commandsElements)
            {
                elem.ReplaceNodes(elem.Descendants("Command").OrderBy(item => item.Attribute("name").Value));
                newRootElement.Add(elem);
            }

            // Add and sort lists
            commandsElements = xmlRootElement.Descendants("Lists");
            foreach (XElement elem in commandsElements)
            {
                elem.ReplaceNodes(elem.Descendants("List").OrderBy(item => item.Attribute("name").Value));
                newRootElement.Add(elem);
            }


            return newRootElement;

        }
        

        private void txtCommandFile1_TextChanged(object sender, EventArgs e)
        {
            UpdateCompareButtonEnableState();
        }

        private void txtCommandFile2_TextChanged(object sender, EventArgs e)
        {
            UpdateCompareButtonEnableState();
        }

        private void UpdateCompareButtonEnableState()
        {
            if (txtCommandFile1.Text.Length > 0 && txtCommandFile2.Text.Length > 0)
            {
                btnCompare.Enabled = true;
            }
            else
            {
                btnCompare.Enabled = false;
            }
        }

        private void cboDiffTool_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDiffToolFieldsEnableState();
        }

        private void UpdateDiffToolFieldsEnableState()
        {
            if (cboDiffTool.Text.Equals(DIFF_TOOL_CUSTOM_NAME))
            {
                txtDiffToolExe.Enabled = true;
                txtDiffToolArgs.Enabled = true;
                btnDiffToolExe.Enabled = true;
                DiffTool diffTool = m_DiffTools[DIFF_TOOL_CUSTOM_NAME];
                txtDiffToolExe.Text = diffTool.Executable;
                txtDiffToolArgs.Text = diffTool.ArgumentTemplate;
            }
            else if (cboDiffTool.Text.Length > 0)
            {
                txtDiffToolExe.Enabled = false;
                txtDiffToolArgs.Enabled = false;
                btnDiffToolExe.Enabled = false;
                DiffTool diffTool = m_DiffTools[cboDiffTool.Text];
                txtDiffToolExe.Text = diffTool.Executable;
                txtDiffToolArgs.Text = diffTool.ArgumentTemplate;
            }
            else
            {
                txtDiffToolExe.Enabled = false;
                txtDiffToolArgs.Enabled = false;
                btnDiffToolExe.Enabled = false;
                txtDiffToolExe.Text = "";
                txtDiffToolArgs.Text = "";
            }
        }

        private void btnDiffToolExe_Click(object sender, EventArgs e)
        {
            HandleExeFileSelect();
        }

        private void mainWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void txtDiffToolExe_TextChanged(object sender, EventArgs e)
        {
            if (cboDiffTool.Text.Equals(DIFF_TOOL_CUSTOM_NAME))
            {
                m_DiffTools[DIFF_TOOL_CUSTOM_NAME].Executable = txtDiffToolExe.Text;
            }
        }

        private void txtDiffToolArgs_TextChanged(object sender, EventArgs e)
        {
            if (cboDiffTool.Text.Equals(DIFF_TOOL_CUSTOM_NAME))
            {
                m_DiffTools[DIFF_TOOL_CUSTOM_NAME].ArgumentTemplate = txtDiffToolArgs.Text;
            }
        }
    }
}
