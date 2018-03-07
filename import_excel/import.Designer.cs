namespace import_excel
{
    partial class import
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
            this.label1 = new System.Windows.Forms.Label();
            this.dbl = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbl = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "数据库";
            // 
            // dbl
            // 
            this.dbl.FormattingEnabled = true;
            this.dbl.Location = new System.Drawing.Point(68, 33);
            this.dbl.Name = "dbl";
            this.dbl.Size = new System.Drawing.Size(121, 20);
            this.dbl.TabIndex = 2;
            this.dbl.TextChanged += new System.EventHandler(this.dbs);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "表格";
            // 
            // tbl
            // 
            this.tbl.FormattingEnabled = true;
            this.tbl.Location = new System.Drawing.Point(68, 74);
            this.tbl.Name = "tbl";
            this.tbl.Size = new System.Drawing.Size(121, 20);
            this.tbl.TabIndex = 4;
            this.tbl.TextChanged += new System.EventHandler(this.tbs);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(23, 122);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "导入";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.imp);
            // 
            // import
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(229, 160);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tbl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dbl);
            this.Controls.Add(this.label1);
            this.Name = "import";
            this.Text = "import";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox dbl;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox tbl;
        private System.Windows.Forms.Button button1;
    }
}