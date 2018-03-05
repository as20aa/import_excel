using System;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace import_excel
{
    public partial class Form2 : Form
    {
        //入口程序，类似于main
        public Form2()
        {
            this.DialogResult = DialogResult.No;
            InitializeComponent();
        }

        //标签
        private void label4_Click(object sender, EventArgs e)
        {

        }
        //标签
        private void label1_Click(object sender, EventArgs e)
        {

        }
        //登陆按钮
        private void log_in(object sender,EventArgs e)
        {
            //新建一个builder，如果直接用Form1.data.builder，对象是不会被创建的
            //设置连接SQL Server的参数并保存在data这个应用程序数据中
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
                //登录失败
                MessageBox.Show("登陆失败！请重试"+ "\r\n"+MessageBoxButtons.OK);
            }
        }
        //登陆取消按钮
        private void login_cancel(object sender,EventArgs e)
        {
            //connection = null;
            this.Close();
        }
    }
}
