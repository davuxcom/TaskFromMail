namespace OutlookAddIn1
{
    partial class SelectTaskForm
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
            this.lsv = new System.Windows.Forms.ListView();
            this.btn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lsv
            // 
            this.lsv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lsv.FullRowSelect = true;
            this.lsv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.lsv.Location = new System.Drawing.Point(12, 12);
            this.lsv.Name = "lsv";
            this.lsv.Size = new System.Drawing.Size(655, 397);
            this.lsv.TabIndex = 1;
            this.lsv.UseCompatibleStateImageBehavior = false;
            // 
            // btn
            // 
            this.btn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn.Location = new System.Drawing.Point(556, 415);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(111, 37);
            this.btn.TabIndex = 2;
            this.btn.Text = "Select";
            this.btn.UseVisualStyleBackColor = true;
            this.btn.Click += new System.EventHandler(this.btn_Click);
            // 
            // SelectTaskForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(679, 459);
            this.Controls.Add(this.btn);
            this.Controls.Add(this.lsv);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "SelectTaskForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Select Task...";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView lsv;
        private System.Windows.Forms.Button btn;
    }
}