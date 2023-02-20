namespace KB_XML_Compare
{
    partial class mainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtCommandFile1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCommandFile1 = new System.Windows.Forms.Button();
            this.btnCommandFile2 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtCommandFile2 = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnCompare = new System.Windows.Forms.Button();
            this.tableLayoutCompare = new System.Windows.Forms.TableLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.cboDiffTool = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtDiffToolExe = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtDiffToolArgs = new System.Windows.Forms.TextBox();
            this.btnDiffToolExe = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.chkDeleteTempFiles = new System.Windows.Forms.CheckBox();
            this.tableLayoutCompare.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtCommandFile1
            // 
            this.txtCommandFile1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCommandFile1.Location = new System.Drawing.Point(119, 13);
            this.txtCommandFile1.Name = "txtCommandFile1";
            this.txtCommandFile1.Size = new System.Drawing.Size(383, 20);
            this.txtCommandFile1.TabIndex = 1;
            this.txtCommandFile1.TextChanged += new System.EventHandler(this.txtCommandFile1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Command File &1:";
            // 
            // btnCommandFile1
            // 
            this.btnCommandFile1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCommandFile1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCommandFile1.Location = new System.Drawing.Point(508, 9);
            this.btnCommandFile1.Name = "btnCommandFile1";
            this.btnCommandFile1.Size = new System.Drawing.Size(31, 23);
            this.btnCommandFile1.TabIndex = 2;
            this.btnCommandFile1.Text = "...";
            this.btnCommandFile1.UseVisualStyleBackColor = true;
            this.btnCommandFile1.Click += new System.EventHandler(this.btnCommandFile1_Click);
            // 
            // btnCommandFile2
            // 
            this.btnCommandFile2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCommandFile2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCommandFile2.Location = new System.Drawing.Point(508, 35);
            this.btnCommandFile2.Name = "btnCommandFile2";
            this.btnCommandFile2.Size = new System.Drawing.Size(31, 23);
            this.btnCommandFile2.TabIndex = 5;
            this.btnCommandFile2.Text = "...";
            this.btnCommandFile2.UseVisualStyleBackColor = true;
            this.btnCommandFile2.Click += new System.EventHandler(this.btnCommandFile2_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(13, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Command File &2:";
            // 
            // txtCommandFile2
            // 
            this.txtCommandFile2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCommandFile2.Location = new System.Drawing.Point(119, 39);
            this.txtCommandFile2.Name = "txtCommandFile2";
            this.txtCommandFile2.Size = new System.Drawing.Size(383, 20);
            this.txtCommandFile2.TabIndex = 4;
            this.txtCommandFile2.TextChanged += new System.EventHandler(this.txtCommandFile2_TextChanged);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // btnCompare
            // 
            this.btnCompare.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnCompare.Location = new System.Drawing.Point(224, 6);
            this.btnCompare.Name = "btnCompare";
            this.btnCompare.Size = new System.Drawing.Size(75, 23);
            this.btnCompare.TabIndex = 0;
            this.btnCompare.Text = "&Compare";
            this.btnCompare.UseVisualStyleBackColor = true;
            this.btnCompare.Click += new System.EventHandler(this.btnCompare_Click);
            // 
            // tableLayoutCompare
            // 
            this.tableLayoutCompare.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutCompare.ColumnCount = 1;
            this.tableLayoutCompare.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutCompare.Controls.Add(this.btnCompare, 0, 0);
            this.tableLayoutCompare.Location = new System.Drawing.Point(16, 275);
            this.tableLayoutCompare.Name = "tableLayoutCompare";
            this.tableLayoutCompare.RowCount = 1;
            this.tableLayoutCompare.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutCompare.Size = new System.Drawing.Size(523, 35);
            this.tableLayoutCompare.TabIndex = 14;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(21, 69);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Diff Tool &Type:";
            // 
            // cboDiffTool
            // 
            this.cboDiffTool.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboDiffTool.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cboDiffTool.FormattingEnabled = true;
            this.cboDiffTool.Location = new System.Drawing.Point(120, 66);
            this.cboDiffTool.Name = "cboDiffTool";
            this.cboDiffTool.Size = new System.Drawing.Size(121, 21);
            this.cboDiffTool.Sorted = true;
            this.cboDiffTool.TabIndex = 7;
            this.cboDiffTool.SelectedIndexChanged += new System.EventHandler(this.cboDiffTool_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(28, 99);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(85, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Diff Tool &Exe:";
            // 
            // txtDiffToolExe
            // 
            this.txtDiffToolExe.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDiffToolExe.Location = new System.Drawing.Point(119, 101);
            this.txtDiffToolExe.Name = "txtDiffToolExe";
            this.txtDiffToolExe.Size = new System.Drawing.Size(383, 20);
            this.txtDiffToolExe.TabIndex = 9;
            this.txtDiffToolExe.TextChanged += new System.EventHandler(this.txtDiffToolExe_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(24, 130);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Diff Tool &Args:";
            // 
            // txtDiffToolArgs
            // 
            this.txtDiffToolArgs.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDiffToolArgs.Location = new System.Drawing.Point(120, 130);
            this.txtDiffToolArgs.Name = "txtDiffToolArgs";
            this.txtDiffToolArgs.Size = new System.Drawing.Size(382, 20);
            this.txtDiffToolArgs.TabIndex = 12;
            this.txtDiffToolArgs.TextChanged += new System.EventHandler(this.txtDiffToolArgs_TextChanged);
            // 
            // btnDiffToolExe
            // 
            this.btnDiffToolExe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDiffToolExe.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDiffToolExe.Location = new System.Drawing.Point(508, 99);
            this.btnDiffToolExe.Name = "btnDiffToolExe";
            this.btnDiffToolExe.Size = new System.Drawing.Size(31, 23);
            this.btnDiffToolExe.TabIndex = 10;
            this.btnDiffToolExe.Text = "...";
            this.btnDiffToolExe.UseVisualStyleBackColor = true;
            this.btnDiffToolExe.Click += new System.EventHandler(this.btnDiffToolExe_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(120, 157);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(393, 13);
            this.label6.TabIndex = 16;
            this.label6.Text = "e.g. \"-file1 {1} -file2 {2}\" where {1} and {2} are placeholders for the command f" +
    "iles.";
            // 
            // chkDeleteTempFiles
            // 
            this.chkDeleteTempFiles.AutoSize = true;
            this.chkDeleteTempFiles.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDeleteTempFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDeleteTempFiles.Location = new System.Drawing.Point(16, 187);
            this.chkDeleteTempFiles.Name = "chkDeleteTempFiles";
            this.chkDeleteTempFiles.Size = new System.Drawing.Size(132, 17);
            this.chkDeleteTempFiles.TabIndex = 13;
            this.chkDeleteTempFiles.Text = "&Delete Temp Files:";
            this.chkDeleteTempFiles.UseVisualStyleBackColor = true;
            // 
            // mainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(551, 322);
            this.Controls.Add(this.chkDeleteTempFiles);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btnDiffToolExe);
            this.Controls.Add(this.txtDiffToolArgs);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtDiffToolExe);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cboDiffTool);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tableLayoutCompare);
            this.Controls.Add(this.btnCommandFile2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCommandFile2);
            this.Controls.Add(this.btnCommandFile1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtCommandFile1);
            this.Name = "mainWindow";
            this.Text = "KB Commands Diff Tool";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.mainWindow_FormClosing);
            this.tableLayoutCompare.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtCommandFile1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCommandFile1;
        private System.Windows.Forms.Button btnCommandFile2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtCommandFile2;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btnCompare;
        private System.Windows.Forms.TableLayoutPanel tableLayoutCompare;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboDiffTool;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtDiffToolExe;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtDiffToolArgs;
        private System.Windows.Forms.Button btnDiffToolExe;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox chkDeleteTempFiles;
    }
}

