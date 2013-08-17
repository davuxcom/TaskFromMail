namespace OutlookAddIn1
{
    partial class SelectFolderForm
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
            this.btn = new System.Windows.Forms.Button();
            this.lsv = new System.Windows.Forms.ListView();
            this.chk = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btn
            // 
            this.btn.Location = new System.Drawing.Point(393, 262);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(111, 37);
            this.btn.TabIndex = 1;
            this.btn.Text = "Select";
            this.btn.UseVisualStyleBackColor = true;
            this.btn.Click += new System.EventHandler(this.btn_Click);
            // 
            // lsv
            // 
            this.lsv.FullRowSelect = true;
            this.lsv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.lsv.HideSelection = false;
            this.lsv.Location = new System.Drawing.Point(12, 12);
            this.lsv.MultiSelect = false;
            this.lsv.Name = "lsv";
            this.lsv.Size = new System.Drawing.Size(492, 244);
            this.lsv.TabIndex = 2;
            this.lsv.UseCompatibleStateImageBehavior = false;
            this.lsv.View = System.Windows.Forms.View.Details;
            // 
            // chk
            // 
            this.chk.AutoSize = true;
            this.chk.Location = new System.Drawing.Point(12, 271);
            this.chk.Name = "chk";
            this.chk.Size = new System.Drawing.Size(196, 21);
            this.chk.TabIndex = 3;
            this.chk.Text = "Always use this task folder";
            this.chk.UseVisualStyleBackColor = true;
            // 
            // SelectFolderForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(516, 309);
            this.Controls.Add(this.chk);
            this.Controls.Add(this.lsv);
            this.Controls.Add(this.btn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "SelectFolderForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create task in folder...";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn;
        private System.Windows.Forms.ListView lsv;
        private System.Windows.Forms.CheckBox chk;

    }
}