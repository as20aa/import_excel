using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using appdata;
using System.Data.SqlClient;

namespace import_excel
{
    public partial class import : Form
    {
        public import()
        {
            
            InitializeComponent();
            dbcheck();
        }
        private void dbcheck()
        {
            //SELECT Name FROM Master..SysDatabases ORDER BY Name;
            if (Form1.data.connection.State != ConnectionState.Open)
             {
                //重新激活连接
                MessageBox.Show("未连接");
              Form1.data.connection.Open();
             }
            try
            {
                //using的作用是创建一个临时的变量和连接，所以当结束using block时内存将会被回收，在这里就算是static变量也只是临时的副本
                using ( SqlDataAdapter adapter = new SqlDataAdapter())
                {
                    string sql;
                    StringBuilder sb = new StringBuilder();
                    sb.Append("select name from ");
                    sb.Append(Form1.data.builder.InitialCatalog);
                    sb.Append("..sysdatabases order by name;");
                    sql = sb.ToString();

                    adapter.SelectCommand = new SqlCommand(sql, Form1.data.connection);
                    DataSet dblist = new DataSet();
                    adapter.Fill(dblist);

                    if (dblist.Tables.Count > 0)
                    {
                        
                        if (dblist.Tables[0].Rows.Count > 0)
                        {
                            //还是不能单独操作datatable
                            foreach (DataRow row in dblist.Tables[0].Rows)
                                foreach (DataColumn col in dblist.Tables[0].Columns)
                                    dbl.Items.Add(row[col].ToString());
                        }
                        else
                        {
                            MessageBox.Show("无数据库");
                        }
                    }
                    else
                    {
                        MessageBox.Show("数据库为空");
                    }
                }

            }
            catch
            {
                MessageBox.Show("查询失败！");
            }
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
    
}
