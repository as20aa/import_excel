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
        //检查数据库并设定数据库
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
                                    //将数据库列表加载到相应的选项框中
                                    dbl.Items.Add(row[col].ToString());
                        }
                    }
                    
                }

            }
            catch
            {
                MessageBox.Show("查询失败！");
            }
            
        }

        //选定数据库
        private void dbs(object sender,EventArgs e)
        {
            //选定数据库
            Form1.data.database = this.dbl.Text;
            tbcheck();
        }

        //获取列表
        private void tbcheck()
        {
            if (Form1.data.connection.State != ConnectionState.Open)
            {
                //重新激活连接
                MessageBox.Show("未连接");
                Form1.data.connection.Open();
            }
            try
            {
                //
                using (SqlDataAdapter adapter = new SqlDataAdapter())
                {
                    SqlCommand sql = new SqlCommand();
                    DataSet temp = new DataSet();
                    //查询所有的表
                    string sb = "use "+Form1.data.database+" select [name] from [sysobjects] where [type] = 'u' order by [name];";
                    adapter.SelectCommand = new SqlCommand(sb, Form1.data.connection);
                    adapter.Fill(temp);
                    if (temp.Tables.Count > 0)
                    {
                        if (temp.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow row in temp.Tables[0].Rows)
                                foreach (DataColumn col in temp.Tables[0].Columns)
                                    tbl.Items.Add(row[col].ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("表查询失败");
                    }
                }
            }
            catch
            {
                MessageBox.Show("表查询失败！");
                this.Close();
            }

        }

        //选定列表
        private void tbs(object sender,EventArgs e)
        {
            Form1.data.table = this.tbl.Text;
        }

        //导入到特定的位置
        private void imp(object sender,EventArgs e)
        {
            try
            { 
                //导入到SQL Server
                //检查数据库
                string sql = "if not exists (select * from sys.databases where name = '"+Form1.data.database+"') create database ["+ Form1.data.database +"]; ";
                using (SqlCommand command = new SqlCommand(sql, Form1.data.connection))
                {
                    command.ExecuteNonQuery();
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("use ");
                sb.Append(Form1.data.database);
                sb.Append(";");
                sb.Append("if not exists (select * from sysobjects where name='");
                sb.Append(Form1.data.table);
                sb.Append("' and xtype='U') create table ");
                sb.Append(Form1.data.table);
                sb.Append(" (");

                //添加列信息
                //据说数据库不能建立空表？？？不然就一个sqldataapdater丢过去
                foreach (DataColumn columns in Form1.data.datatable.Columns)
                {
                    sb.Append(" ");
                    //插入的数据的列名称
                    sb.Append(columns.ColumnName.ToString());
                    sb.Append(" ");
                    //插入列的数据类型
                    if (columns.DataType == System.Type.GetType("System.String"))
                        sb.Append("nvarchar(50)");
                    if (columns.DataType == System.Type.GetType("System.Int64"))
                        sb.Append("bigint");
                    else
                    {
                        if (columns.DataType == System.Type.GetType("System.Int32"))
                            sb.Append("int");
                        else
                        {
                            if (columns.DataType == System.Type.GetType("System.Int16"))
                                sb.Append("smallint");
                            else
                                if (columns.DataType == System.Type.GetType("System.Double"))
                                sb.Append("real");
                        }
                    }
                    if (columns.DataType == System.Type.GetType("System.DateTime"))
                        sb.Append("datetime");
                    sb.Append(",");
                }
                //移除最后一个','
                sb = sb.Remove(sb.Length - 1, 1);
                sb.Append(");");
                sql = sb.ToString();

                //创建SQL数据库操作
                using (SqlCommand command = new SqlCommand(sql, Form1.data.connection))
                {
                    command.ExecuteNonQuery();
                    //Console.WriteLine("Table creation is done.");
                }

                //从datatable的每一行开始筛选，对记录进行查重
                foreach (DataRow row2 in Form1.data.datatable.Rows)
                {
                    int repeat = 0;//每一行循环均会重设
                    foreach(DataColumn col2 in Form1.data.datatable.Columns)
                    {
                        StringBuilder jd = new StringBuilder();
                        //向SQL Server发出查询指令，并根据返回数据判断是否存在重复记录
                        jd.Append("select * from ");
                        jd.Append(Form1.data.database);
                        jd.Append("..");
                        jd.Append(Form1.data.table);
                        jd.Append(" where ");
                        jd.Append(col2.ColumnName);//列的名字
                        if (row2[col2.ColumnName] is System.DBNull)
                        {
                            jd.Append(" is null;"); 
                        }
                        else
                        {
                            jd.Append(" = '");
                            jd.Append(row2[col2.ColumnName]);
                            jd.Append("';");
                        }
                        sql = jd.ToString();
                        DataSet dataset = new DataSet();
                        //SqlDataAdapter类的用法，接收到的数据是DataSet或者是DataTable类型

                        SqlDataAdapter adapter = new SqlDataAdapter(sql, Form1.data.connection);
                        //将接收到的数据填充到dataset中
                        adapter.Fill(dataset);
                        //采用以下的方式才是正确的对datatable和dataset单元数据的读取方式
                        //MessageBox.Show(temp.Rows[0][0].ToString());
                        //MessageBox.Show(dataset.Tables[0].Rows[0][0].ToString());
                        //adapter.SelectCommand = new SqlCommand(sql, data.connection);
                        //使用dataset中表的行数判断查询结果是否为空
                        if (dataset.Tables[0].Rows.Count > 0)
                        {

                            //一旦dataset为空，则下面的判断语句就会出错！
                            if (row2[0].ToString() == dataset.Tables[0].Rows[0][0].ToString())
                            {
                                repeat++;
                                //弹出信息框，对信息框方法的描述见定义
                                //注意添加的按钮属性是okcancel而不是yesno
                            }
                        }//if
                    }
                    //判断重复的情况
                    if(repeat==Form1.data.datatable.Columns.Count&& Form1.data.datatable.Columns.Count!=0)
                    {
                        MessageBox.Show("表格完全重复！将插入下一行或结束");
                    }
                    else
                    {
                        //对部分重复的情况
                        if (repeat > 0)
                        {
                            DialogResult dr = MessageBox.Show("表格部分重复！是否更新该列表？", "警告", MessageBoxButtons.OKCancel);
                            if (dr == DialogResult.OK)
                            {
                                //更新该列表
                                StringBuilder di = new StringBuilder();
                                di.Append("delete from ");
                                di.Append(Form1.data.table);
                                di.Append(" where ");
                                di.Append(Form1.data.datatable.Columns[0].ColumnName);
                                di.Append(" = ");
                                di.Append(row2[Form1.data.datatable.Columns[0]].ToString());
                                di.Append(";");
                                sql = di.ToString();
                                using (SqlCommand command2 = new SqlCommand(sql, Form1.data.connection))
                                {
                                    command2.ExecuteNonQuery();
                                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(Form1.data.connection))
                                    {
                                        bulkCopy.DestinationTableName = "dbo." + Form1.data.table;
                                        // Write from the source to the destination.
                                        try
                                        {
                                            //从当前的datatable中克隆到新的datatable中，克隆得到的datatable具有原来
                                            //datatable的架构，但是没有数据，即只有列，没有行数据
                                            DataTable buff = Form1.data.datatable.Clone();
                                            //不能用buff.rows.add(row2),程序会直接出错，可以中buff.importrow(row2)来实现导入一行
                                            buff.ImportRow(row2);
                                            // Write from the source to the destination.
                                            bulkCopy.WriteToServer(buff);
                                            //MessageBox.Show("写入成功！");
                                            //由于buff是局部变量，在每次调用之后会被自动清除故可以不用移除操作
                                            //buff.Rows.Remove(row2);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                        }
                                    }
                                }
                            }

                        }

                        //完全不重复的情况
                        else
                        {
                            
                            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(Form1.data.connection))
                            {
                                bulkCopy.DestinationTableName = "dbo." + Form1.data.table;
                                // Write from the source to the destination.
                                try
                                {
                                    //从当前的datatable中克隆到新的datatable中，克隆得到的datatable具有原来
                                    //datatable的架构，但是没有数据，即只有列，没有行数据
                                    DataTable buff = Form1.data.datatable.Clone();
                                    //不能用buff.rows.add(row2),程序会直接出错，可以中buff.importrow(row2)来实现导入一行
                                    buff.ImportRow(row2);
                                    // Write from the source to the destination.
                                    bulkCopy.WriteToServer(buff);
                                    //MessageBox.Show("写入成功！");
                                    //由于buff是局部变量，在每次调用之后会被自动清除故可以不用移除操作
                                    //buff.Rows.Remove(row2);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                            }
                        }
                    }
                }
                MessageBox.Show("操作完成！");
            }
            catch
            {
                MessageBox.Show("写入失败!");
            }
        }

        private void cancel(object sender,EventArgs e)
        {
            this.Close();
        }

    }
    
}
