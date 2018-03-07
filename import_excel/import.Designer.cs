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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(import));
            this.label1 = new System.Windows.Forms.Label();
            this.dbl = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbl = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.canc = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // dbl
            // 
            this.dbl.FormattingEnabled = true;
            resources.ApplyResources(this.dbl, "dbl");
            this.dbl.Name = "dbl";
            this.dbl.TextChanged += new System.EventHandler(this.dbs);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // tbl
            // 
            this.tbl.FormattingEnabled = true;
            resources.ApplyResources(this.tbl, "tbl");
            this.tbl.Name = "tbl";
            this.tbl.TextChanged += new System.EventHandler(this.tbs);
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.imp);
            // 
            // canc
            // 
            resources.ApplyResources(this.canc, "canc");
            this.canc.Name = "canc";
            this.canc.UseVisualStyleBackColor = true;
            this.canc.Click += new System.EventHandler(this.cancel);
            // 
            // import
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.canc);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tbl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dbl);
            this.Controls.Add(this.label1);
            this.Name = "import";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox dbl;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox tbl;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button canc;
    }
}