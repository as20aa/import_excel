namespace import_excel
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.fileselect = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.itd = new System.Windows.Forms.Button();
            this.its = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // fileselect
            // 
            this.fileselect.Location = new System.Drawing.Point(271, 47);
            this.fileselect.Name = "fileselect";
            this.fileselect.Size = new System.Drawing.Size(112, 23);
            this.fileselect.TabIndex = 0;
            this.fileselect.Text = "浏览";
            this.fileselect.UseVisualStyleBackColor = true;
            this.fileselect.Click += new System.EventHandler(this.filedialog);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(23, 37);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(220, 209);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.Click += new System.EventHandler(this.itdatatable);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 276);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "未选择文件";
            // 
            // itd
            // 
            this.itd.Location = new System.Drawing.Point(271, 98);
            this.itd.Name = "itd";
            this.itd.Size = new System.Drawing.Size(112, 23);
            this.itd.TabIndex = 3;
            this.itd.Text = "导入到datatable";
            this.itd.UseVisualStyleBackColor = true;
            this.itd.MouseClick += new System.Windows.Forms.MouseEventHandler(this.itdatatable);
            // 
            // its
            // 
            this.its.Location = new System.Drawing.Point(271, 150);
            this.its.Name = "its";
            this.its.Size = new System.Drawing.Size(112, 23);
            this.its.TabIndex = 4;
            this.its.Text = "导入SQL Servver";
            this.its.UseVisualStyleBackColor = true;
            this.its.Click += new System.EventHandler(this.itsql);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(201, 275);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "未连接数据库";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(395, 316);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.its);
            this.Controls.Add(this.itd);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.fileselect);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        
        private System.Windows.Forms.Button fileselect;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button itd;
        private System.Windows.Forms.Button its;
        private System.Windows.Forms.Label label2;
    }
}

