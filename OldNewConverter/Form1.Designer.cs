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
            this.stationsDataGridView = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.stationsDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // startStopButton
            // 
            this.startStopButton.Location = new System.Drawing.Point(110, 12);
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
            this.statusLabel.ForeColor = System.Drawing.Color.Red;
            this.statusLabel.Location = new System.Drawing.Point(57, 18);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(47, 13);
            this.statusLabel.TabIndex = 1;
            this.statusLabel.Text = "Stopped";
            // 
            // stationsDataGridView
            // 
            this.stationsDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.stationsDataGridView.Location = new System.Drawing.Point(22, 51);
            this.stationsDataGridView.Name = "stationsDataGridView";
            this.stationsDataGridView.Size = new System.Drawing.Size(798, 198);
            this.stationsDataGridView.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Status:";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(844, 261);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.stationsDataGridView);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.startStopButton);
            this.MaximumSize = new System.Drawing.Size(860, 300);
            this.MinimumSize = new System.Drawing.Size(860, 300);
            this.Name = "Form1";
            this.Text = "Old New Converter";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.stationsDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button startStopButton;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.DataGridView stationsDataGridView;
        private System.Windows.Forms.Label label1;
    }
}

