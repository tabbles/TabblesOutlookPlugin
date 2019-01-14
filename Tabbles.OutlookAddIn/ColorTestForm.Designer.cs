namespace Tabbles.OutlookAddIn
{
    partial class ColorTestForm
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
            this.txtInput = new System.Windows.Forms.TextBox();
            this.lblResult = new System.Windows.Forms.Label();
            this.btnEval = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtInput
            // 
            this.txtInput.Location = new System.Drawing.Point(12, 32);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(100, 20);
            this.txtInput.TabIndex = 0;
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Location = new System.Drawing.Point(12, 108);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(10, 13);
            this.lblResult.TabIndex = 1;
            this.lblResult.Text = "-";
            // 
            // btnEval
            // 
            this.btnEval.Location = new System.Drawing.Point(170, 29);
            this.btnEval.Name = "btnEval";
            this.btnEval.Size = new System.Drawing.Size(75, 23);
            this.btnEval.TabIndex = 2;
            this.btnEval.Text = "Evaluate";
            this.btnEval.UseVisualStyleBackColor = true;
            this.btnEval.Click += new System.EventHandler(this.btnEval_Click);
            // 
            // ColorTestForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(300, 130);
            this.Controls.Add(this.btnEval);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.txtInput);
            this.Name = "ColorTestForm";
            this.Text = "ColorTestForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtInput;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Button btnEval;
    }
}