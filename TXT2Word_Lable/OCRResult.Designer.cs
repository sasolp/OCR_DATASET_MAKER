namespace TXT2Word_Lable
{
    partial class OCRResult
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
            this.lblOCR = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblOCR
            // 
            this.lblOCR.AutoSize = true;
            this.lblOCR.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblOCR.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.lblOCR.Location = new System.Drawing.Point(0, 0);
            this.lblOCR.Name = "lblOCR";
            this.lblOCR.Size = new System.Drawing.Size(35, 13);
            this.lblOCR.TabIndex = 0;
            this.lblOCR.Text = "label1";
            this.lblOCR.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // OCRResult
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 470);
            this.Controls.Add(this.lblOCR);
            this.Name = "OCRResult";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "OCRResult";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label lblOCR;
    }
}