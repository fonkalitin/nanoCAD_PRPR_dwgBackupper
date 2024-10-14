namespace prPr_FileBackupper
{
    partial class MainForm
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
            this.revComBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // revComBox
            // 
            this.revComBox.FormattingEnabled = true;
            this.revComBox.Items.AddRange(new object[] {
            "Rev.0",
            "Rev.1",
            "Rev.2",
            "Rev.3",
            "Rev.4",
            "Rev.5",
            "Rev.6",
            "Rev.7",
            "Rev.8",
            "Rev.9",
            "Rev.10"});
            this.revComBox.Location = new System.Drawing.Point(136, 105);
            this.revComBox.Name = "revComBox";
            this.revComBox.Size = new System.Drawing.Size(139, 27);
            this.revComBox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(101, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(205, 19);
            this.label1.TabIndex = 1;
            this.label1.Text = "Укажите обозначение ревизии";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(134, 171);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(141, 26);
            this.button1.TabIndex = 2;
            this.button1.Text = "Создать ревизию";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(421, 370);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.revComBox);
            this.Name = "MainForm";
            this.Text = "Резервное копирование";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ComboBox revComBox;
        private Label label1;
        private Button button1;
    }
}