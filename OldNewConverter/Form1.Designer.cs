namespace OldNewConverter
{
    partial class Form1
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
            this.readTxtFileButton = new System.Windows.Forms.Button();
            this.configPathTextBox = new System.Windows.Forms.TextBox();
            this.readExcelFileButton = new System.Windows.Forms.Button();
            this.originFolderLabel = new System.Windows.Forms.Label();
            this.originFolderTextBox = new System.Windows.Forms.TextBox();
            this.destinationFolderTextBox = new System.Windows.Forms.TextBox();
            this.destinationFolderLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // readTxtFileButton
            // 
            this.readTxtFileButton.Location = new System.Drawing.Point(413, 86);
            this.readTxtFileButton.Name = "readTxtFileButton";
            this.readTxtFileButton.Size = new System.Drawing.Size(75, 23);
            this.readTxtFileButton.TabIndex = 0;
            this.readTxtFileButton.Text = "Read txt";
            this.readTxtFileButton.UseVisualStyleBackColor = true;
            this.readTxtFileButton.Click += new System.EventHandler(this.readTxtFileButton_Click);
            // 
            // configPathTextBox
            // 
            this.configPathTextBox.Location = new System.Drawing.Point(12, 12);
            this.configPathTextBox.Name = "configPathTextBox";
            this.configPathTextBox.Size = new System.Drawing.Size(224, 20);
            this.configPathTextBox.TabIndex = 1;
            this.configPathTextBox.Text = "C:\\OldNewGateway\\config\\stations.xlsx";
            // 
            // readExcelFileButton
            // 
            this.readExcelFileButton.AllowDrop = true;
            this.readExcelFileButton.Location = new System.Drawing.Point(242, 10);
            this.readExcelFileButton.Name = "readExcelFileButton";
            this.readExcelFileButton.Size = new System.Drawing.Size(75, 23);
            this.readExcelFileButton.TabIndex = 2;
            this.readExcelFileButton.Text = "Read Excel File";
            this.readExcelFileButton.UseVisualStyleBackColor = true;
            this.readExcelFileButton.Click += new System.EventHandler(this.readExcelFileButton_Click);
            // 
            // originFolderLabel
            // 
            this.originFolderLabel.AutoSize = true;
            this.originFolderLabel.Location = new System.Drawing.Point(39, 54);
            this.originFolderLabel.Name = "originFolderLabel";
            this.originFolderLabel.Size = new System.Drawing.Size(100, 13);
            this.originFolderLabel.TabIndex = 3;
            this.originFolderLabel.Text = "Origin Folder Name:";
            // 
            // originFolderTextBox
            // 
            this.originFolderTextBox.Location = new System.Drawing.Point(145, 51);
            this.originFolderTextBox.Name = "originFolderTextBox";
            this.originFolderTextBox.Size = new System.Drawing.Size(229, 20);
            this.originFolderTextBox.TabIndex = 4;
            // 
            // destinationFolderTextBox
            // 
            this.destinationFolderTextBox.Location = new System.Drawing.Point(145, 78);
            this.destinationFolderTextBox.Name = "destinationFolderTextBox";
            this.destinationFolderTextBox.Size = new System.Drawing.Size(229, 20);
            this.destinationFolderTextBox.TabIndex = 6;
            // 
            // destinationFolderLabel
            // 
            this.destinationFolderLabel.AutoSize = true;
            this.destinationFolderLabel.Location = new System.Drawing.Point(13, 81);
            this.destinationFolderLabel.Name = "destinationFolderLabel";
            this.destinationFolderLabel.Size = new System.Drawing.Size(126, 13);
            this.destinationFolderLabel.TabIndex = 5;
            this.destinationFolderLabel.Text = "Destination Folder Name:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 121);
            this.Controls.Add(this.destinationFolderTextBox);
            this.Controls.Add(this.destinationFolderLabel);
            this.Controls.Add(this.originFolderTextBox);
            this.Controls.Add(this.originFolderLabel);
            this.Controls.Add(this.readExcelFileButton);
            this.Controls.Add(this.configPathTextBox);
            this.Controls.Add(this.readTxtFileButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button readTxtFileButton;
        private System.Windows.Forms.TextBox configPathTextBox;
        private System.Windows.Forms.Button readExcelFileButton;
        private System.Windows.Forms.Label originFolderLabel;
        private System.Windows.Forms.TextBox originFolderTextBox;
        private System.Windows.Forms.TextBox destinationFolderTextBox;
        private System.Windows.Forms.Label destinationFolderLabel;
    }
}

