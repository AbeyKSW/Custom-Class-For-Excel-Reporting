namespace PerformanceViewer
{
    partial class PerformanceViewer
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
            this.cmbDivCode = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dtFrom = new System.Windows.Forms.DateTimePicker();
            this.dtTo = new System.Windows.Forms.DateTimePicker();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.prBar = new System.Windows.Forms.ProgressBar();
            this.prBar1 = new System.Windows.Forms.ProgressBar();
            this.prBar2 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // cmbDivCode
            // 
            this.cmbDivCode.FormattingEnabled = true;
            this.cmbDivCode.Items.AddRange(new object[] {
            "15",
            "16",
            "17",
            "18",
            "19"});
            this.cmbDivCode.Location = new System.Drawing.Point(106, 27);
            this.cmbDivCode.Name = "cmbDivCode";
            this.cmbDivCode.Size = new System.Drawing.Size(121, 21);
            this.cmbDivCode.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Division Code";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "From Date";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 115);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "To Date";
            // 
            // dtFrom
            // 
            this.dtFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtFrom.Location = new System.Drawing.Point(106, 72);
            this.dtFrom.Name = "dtFrom";
            this.dtFrom.Size = new System.Drawing.Size(121, 20);
            this.dtFrom.TabIndex = 4;
            // 
            // dtTo
            // 
            this.dtTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtTo.Location = new System.Drawing.Point(106, 115);
            this.dtTo.Name = "dtTo";
            this.dtTo.Size = new System.Drawing.Size(121, 20);
            this.dtTo.TabIndex = 5;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(152, 150);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 39);
            this.btnGenerate.TabIndex = 6;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // prBar
            // 
            this.prBar.Location = new System.Drawing.Point(30, 234);
            this.prBar.Name = "prBar";
            this.prBar.Size = new System.Drawing.Size(197, 12);
            this.prBar.TabIndex = 7;
            // 
            // prBar1
            // 
            this.prBar1.Location = new System.Drawing.Point(30, 216);
            this.prBar1.Name = "prBar1";
            this.prBar1.Size = new System.Drawing.Size(197, 12);
            this.prBar1.TabIndex = 8;
            // 
            // prBar2
            // 
            this.prBar2.Location = new System.Drawing.Point(30, 198);
            this.prBar2.Name = "prBar2";
            this.prBar2.Size = new System.Drawing.Size(197, 12);
            this.prBar2.TabIndex = 9;
            // 
            // PerformanceViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(259, 258);
            this.Controls.Add(this.prBar2);
            this.Controls.Add(this.prBar1);
            this.Controls.Add(this.prBar);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.dtTo);
            this.Controls.Add(this.dtFrom);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbDivCode);
            this.Name = "PerformanceViewer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Performance Viewer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbDivCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtFrom;
        private System.Windows.Forms.DateTimePicker dtTo;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.ProgressBar prBar;
        private System.Windows.Forms.ProgressBar prBar1;
        private System.Windows.Forms.ProgressBar prBar2;
    }
}

