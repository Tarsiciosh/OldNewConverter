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
            this.startStopButton = new System.Windows.Forms.Button();
            this.statusLabel = new System.Windows.Forms.Label();
            this.originLabel = new System.Windows.Forms.Label();
            this.DestinationLabel = new System.Windows.Forms.Label();
            this.originTextBox = new System.Windows.Forms.TextBox();
            this.destinationTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // startStopButton
            // 
            this.startStopButton.Location = new System.Drawing.Point(12, 12);
            this.startStopButton.Name = "startStopButton";
            this.startStopButton.Size = new System.Drawing.Size(71, 24);
            this.startStopButton.TabIndex = 0;
            this.startStopButton.Text = "Start";
            this.startStopButton.UseVisualStyleBackColor = true;
            this.startStopButton.Click += new System.EventHandler(this.startStopButton_Click);
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(106, 18);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(45, 13);
            this.statusLabel.TabIndex = 1;
            this.statusLabel.Text = "stopped";
            // 
            // originLabel
            // 
            this.originLabel.AutoSize = true;
            this.originLabel.Location = new System.Drawing.Point(57, 83);
            this.originLabel.Name = "originLabel";
            this.originLabel.Size = new System.Drawing.Size(81, 13);
            this.originLabel.TabIndex = 2;
            this.originLabel.Text = "Origin File Path:";
            // 
            // DestinationLabel
            // 
            this.DestinationLabel.AutoSize = true;
            this.DestinationLabel.Location = new System.Drawing.Point(31, 117);
            this.DestinationLabel.Name = "DestinationLabel";
            this.DestinationLabel.Size = new System.Drawing.Size(107, 13);
            this.DestinationLabel.TabIndex = 3;
            this.DestinationLabel.Text = "Destination File Path:";
            // 
            // originTextBox
            // 
            this.originTextBox.Location = new System.Drawing.Point(144, 80);
            this.originTextBox.Name = "originTextBox";
            this.originTextBox.Size = new System.Drawing.Size(300, 20);
            this.originTextBox.TabIndex = 4;
            this.originTextBox.Text = "C:\\FTP Server Results\\Origin";
            // 
            // destinationTextBox
            // 
            this.destinationTextBox.Location = new System.Drawing.Point(144, 114);
            this.destinationTextBox.Name = "destinationTextBox";
            this.destinationTextBox.Size = new System.Drawing.Size(300, 20);
            this.destinationTextBox.TabIndex = 5;
            this.destinationTextBox.Text = "C:\\FTP Server Results\\Destination";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(520, 190);
            this.Controls.Add(this.destinationTextBox);
            this.Controls.Add(this.originTextBox);
            this.Controls.Add(this.DestinationLabel);
            this.Controls.Add(this.originLabel);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.startStopButton);
            this.Name = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button startStopButton;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Label originLabel;
        private System.Windows.Forms.Label DestinationLabel;
        private System.Windows.Forms.TextBox originTextBox;
        private System.Windows.Forms.TextBox destinationTextBox;
    }
}

