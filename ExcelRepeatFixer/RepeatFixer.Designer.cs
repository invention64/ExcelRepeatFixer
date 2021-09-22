
namespace ExcelRepeatFixer
{
    partial class RepeatFixer
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
            this.btnBrowseFiles = new System.Windows.Forms.Button();
            this.fileSelectedTxt = new System.Windows.Forms.Label();
            this.backupChk = new System.Windows.Forms.CheckBox();
            this.fieldCombo = new System.Windows.Forms.ComboBox();
            this.fieldLabel = new System.Windows.Forms.Label();
            this.btnProcess = new System.Windows.Forms.Button();
            this.openExcelFile = new System.Windows.Forms.OpenFileDialog();
            this.lblError = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnBrowseFiles
            // 
            this.btnBrowseFiles.Location = new System.Drawing.Point(35, 31);
            this.btnBrowseFiles.Name = "btnBrowseFiles";
            this.btnBrowseFiles.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseFiles.TabIndex = 0;
            this.btnBrowseFiles.Text = "Browse Files";
            this.btnBrowseFiles.UseVisualStyleBackColor = true;
            this.btnBrowseFiles.Click += new System.EventHandler(this.btnBrowseFiles_Click);
            // 
            // fileSelectedTxt
            // 
            this.fileSelectedTxt.AutoSize = true;
            this.fileSelectedTxt.Location = new System.Drawing.Point(116, 36);
            this.fileSelectedTxt.Name = "fileSelectedTxt";
            this.fileSelectedTxt.Size = new System.Drawing.Size(85, 13);
            this.fileSelectedTxt.TabIndex = 1;
            this.fileSelectedTxt.Text = "No File Selected";
            // 
            // backupChk
            // 
            this.backupChk.AutoSize = true;
            this.backupChk.Location = new System.Drawing.Point(35, 69);
            this.backupChk.Name = "backupChk";
            this.backupChk.Size = new System.Drawing.Size(93, 17);
            this.backupChk.TabIndex = 2;
            this.backupChk.Text = "Make Backup";
            this.backupChk.UseVisualStyleBackColor = true;
            // 
            // fieldCombo
            // 
            this.fieldCombo.FormattingEnabled = true;
            this.fieldCombo.Location = new System.Drawing.Point(119, 92);
            this.fieldCombo.Name = "fieldCombo";
            this.fieldCombo.Size = new System.Drawing.Size(121, 21);
            this.fieldCombo.TabIndex = 3;
            // 
            // fieldLabel
            // 
            this.fieldLabel.AutoSize = true;
            this.fieldLabel.Location = new System.Drawing.Point(32, 95);
            this.fieldLabel.Name = "fieldLabel";
            this.fieldLabel.Size = new System.Drawing.Size(78, 13);
            this.fieldLabel.TabIndex = 4;
            this.fieldLabel.Text = "Field to Check:";
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(98, 119);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(75, 23);
            this.btnProcess.TabIndex = 5;
            this.btnProcess.Text = "Process";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // openExcelFile
            // 
            this.openExcelFile.DefaultExt = "xlsx";
            this.openExcelFile.Filter = "Excel Files|*.xlsx";
            this.openExcelFile.FileOk += new System.ComponentModel.CancelEventHandler(this.openExcelFile_FileOk);
            // 
            // lblError
            // 
            this.lblError.AutoSize = true;
            this.lblError.Font = new System.Drawing.Font("Microsoft PhagsPa", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblError.ForeColor = System.Drawing.Color.Red;
            this.lblError.Location = new System.Drawing.Point(104, 145);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(52, 24);
            this.lblError.TabIndex = 6;
            this.lblError.Text = "Error";
            this.lblError.Visible = false;
            // 
            // RepeatFixer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(304, 189);
            this.Controls.Add(this.lblError);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.fieldLabel);
            this.Controls.Add(this.fieldCombo);
            this.Controls.Add(this.backupChk);
            this.Controls.Add(this.fileSelectedTxt);
            this.Controls.Add(this.btnBrowseFiles);
            this.Name = "RepeatFixer";
            this.Text = "Excel Repeat Fixer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowseFiles;
        private System.Windows.Forms.Label fileSelectedTxt;
        private System.Windows.Forms.CheckBox backupChk;
        private System.Windows.Forms.ComboBox fieldCombo;
        private System.Windows.Forms.Label fieldLabel;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.OpenFileDialog openExcelFile;
        private System.Windows.Forms.Label lblError;
    }
}

