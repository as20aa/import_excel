using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using appdata;

namespace import_excel
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            this.DialogResult = DialogResult.No;
            InitializeComponent();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void log_in(object sender,EventArgs e)
        {
            Form1.data.builder=new SqlConnectionStringBuilder();
            Form1.data.builder.DataSource = textBox1.Text.ToString();
            Form1.data.builder.UserID = textBox2.Text.ToString();
            Form1.data.builder.Password = textBox3.Text.ToString();
            Form1.data.builder.InitialCatalog = textBox4.Text.ToString();

            //connect to sql server
            try
            {
                using (Form1.data.connection = new SqlConnection(Form1.data.builder.ConnectionString))
                {
                    //登陆成功
                    Form1.data.connection.Open();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch
            {
                MessageBox.Show("登陆失败！请重试"+ "\r\n"+MessageBoxButtons.OK);
            }
        }

        private void login_cancel(object sender,EventArgs e)
        {
            //connection = null;
            this.Close();
        }
    }
}
