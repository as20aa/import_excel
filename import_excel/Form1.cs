using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using appdata;

namespace import_excel
{
    public partial class Form1 : Form
    {
        public static appdatas data;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //窗口加载之后执行的代码
            data = new appdatas();
            Form2 loginw = new Form2();
            if (loginw.ShowDialog() != DialogResult.OK)
            {
                this.Close();
            }
            label2.Text = "已连接数据库";
        }
        private void filedialog(object sender,EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Multiselect = true;
            file.Title = "请选择文件";
            file.Filter = "所有文件(*.*)|*.*|excel 2013及更高(*.xlsx*)|*.xlsx*|excel 2013以前(*.xls*)|*.xls*";
            if (file.ShowDialog() == DialogResult.OK)
            {
                data.path = file.FileName;
                label1.Text = "已选择文件";
                //MessageBox.Show("已选择文件:" + path, "选择文件提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void itdatatable(object sender,EventArgs e)
        {
            //用excel com打开目标文件
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(data.path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            data.datatable = new DataTable("datatable");
            //data.dtt=new DataTable("dtt");
            //添加列
            DataColumn column;
            for (int i = 1; i <= colCount; i++)
            {
                if (xlRange.Cells[1, i].Value2 != null && xlRange.Cells[1, i] != null)
                {
                    column = new DataColumn();
                    column.ColumnName = xlRange.Cells[1, i].Value2.ToString();
                    if (xlRange.Cells[2, i].value != null)
                    {
                        //对datatime型数据要特别处理
                        if (xlRange.Cells[2, i].value is DateTime)
                            column.DataType = System.Type.GetType("System.DateTime");
                        else
                            column.DataType = xlRange.Cells[2, i].Value2.GetType();
                    }
                    else
                        column.DataType = System.Type.GetType("System.String");
                    column.ReadOnly = false;
                    column.Unique = false;
                    data.datatable.Columns.Add(column);
                }
            }
            //此处的dtt是接收了datatable的列信息，但是两个实际上是绑定在一起了，没有
            //data.dtt = data.datatable;

            //添加行
            DataRow row;
            for (int i = 2; i <= rowCount; i++)
            {
                row = data.datatable.NewRow();
                for (int j = 0; j <= colCount - 1; j++)
                {
                    //cells从1开始计数
                    if (xlRange.Cells[i, j + 1].Value2 != null)
                    {
                        //datatime型数据另外处理
                        if (xlRange.Cells[i, j + 1].value is DateTime)
                        {
                            string strValue = xlRange.Cells[i, j + 1].Value2.ToString(); //获取得到数字值
                            //注意数据表中含有的日期数据精确到了小时，所以表示日期应该用double，而不是表示日的int32
                            string strDate = DateTime.FromOADate(Convert.ToDouble(strValue)).ToString("s");
                            ////转成sql server能接受的数据格式
                            row[data.datatable.Columns[j].ColumnName] = strDate;
                        }
                        //将相应列的数据导入到datatable中，注意位置的对应关系，在column中从0开始计数
                        else
                        {
                            row[data.datatable.Columns[j].ColumnName] = xlRange.Cells[i, j + 1].Value2;
                        }
                    }
                }
                data.datatable.Rows.Add(row);
            }

            //在datagridview中显示数据
            this.dataGridView1.DataSource = data.datatable;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        private void itsql(object sender, EventArgs e)
        {
            try
            {
                using (data.connection = new SqlConnection(data.builder.ConnectionString))
                {
                    //登陆成功
                    data.connection.Open();
                    //导入到SQL Server
                    //检查数据库
                    string sql = "if not exists (select * from sys.databases where name = 'list') create database [list];";
                    using (SqlCommand command = new SqlCommand(sql, data.connection))
                    {
                        command.ExecuteNonQuery();
                        //Console.WriteLine("Database check and creation is done.");
                    }
                    StringBuilder sb = new StringBuilder();
                    sb.Append("use list;");
                    sb.Append("if not exists (select * from sysobjects where name='datatable' and xtype='U') create table datatable (");

                    //添加列信息
                    foreach (DataColumn columns in data.datatable.Columns)
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
                    using (SqlCommand command = new SqlCommand(sql, data.connection))
                    {
                        command.ExecuteNonQuery();
                        //Console.WriteLine("Table creation is done.");
                    }

                    //从datatable的每一行开始筛选，对重复自定义码数据进行操作

                    foreach (DataRow row2 in data.datatable.Rows)
                    {
                        StringBuilder jd = new StringBuilder();
                        //如果用select语句，则如果不存在该行则会直接退出程序
                        jd.Append("select * from datatable where ");
                        jd.Append(data.datatable.Columns[0].ColumnName);
                        jd.Append(" = ");
                        jd.Append(row2[data.datatable.Columns[0].ColumnName]);
                        jd.Append(";");
                        sql = jd.ToString();
                        //Console.WriteLine(sql);
                        DataSet dataset = new DataSet();
                        SqlDataAdapter adapter = new SqlDataAdapter(sql, data.connection);

                        adapter.Fill(dataset);
                        //采用以下的方式才是正确的读取方式
                        //MessageBox.Show(temp.Rows[0][0].ToString());
                        //MessageBox.Show(dataset.Tables[0].Rows[0][0].ToString());
                        //adapter.SelectCommand = new SqlCommand(sql, data.connection);
                        //此处的判断条件不好！如果找不到数据selectcommand不为空
                        //使用dataset中表的行数判断查询结果是否为空
                        if (dataset.Tables[0].Rows.Count > 0)
                        {

                            //一旦dataset为空，则下面的判断语句就会出错！
                            if (row2[0].ToString() == dataset.Tables[0].Rows[0][0].ToString())
                            {
                                DialogResult dr = MessageBox.Show("表格重复！是否更新该列表？", "警告", MessageBoxButtons.OKCancel);
                                if (dr == DialogResult.OK)
                                {
                                    //更新该列表
                                    StringBuilder di = new StringBuilder();
                                    di.Append("delete from datatable where ");
                                    di.Append(data.datatable.Columns[0].ColumnName);
                                    di.Append(" = ");
                                    di.Append(row2[data.datatable.Columns[0]].ToString());
                                    di.Append(";");
                                    sql = di.ToString();
                                    using (SqlCommand command2 = new SqlCommand(sql, data.connection))
                                    {
                                        command2.ExecuteNonQuery();
                                        //Console.WriteLine("Delete completed");
                                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(data.connection))
                                        {
                                            bulkCopy.DestinationTableName = "dbo.datatable";
                                            // Write from the source to the destination.
                                            try
                                            {
                                                DataTable buff = data.datatable.Clone();
                                                //此处添加行出了问题,不能用buff.rows.add(row2),程序会直接出错，可以中buff.importrow(row2)来实现导入一行
                                                buff.ImportRow(row2);
                                                // Write from the source to the destination.
                                                bulkCopy.WriteToServer(buff);
                                                buff.Rows.Remove(row2);
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine(ex.Message);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(data.connection))
                            {
                                bulkCopy.DestinationTableName = "dbo.datatable";
                                try
                                {
                                    DataTable buff = data.datatable.Clone();
                                    //此处添加行出了问题,不能用buff.rows.add(row2),程序会直接出错，可以中buff.importrow(row2)来实现导入一行
                                    buff.ImportRow(row2);
                                    // Write from the source to the destination.
                                    bulkCopy.WriteToServer(buff);
                                    buff.Rows.Remove(row2);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                            }
                        }

                    }
                    //Console.WriteLine("Done.");
                }
            }
            catch
            {
                label2.Text = "数据库未连接";
                MessageBox.Show("登陆失败！请重试" + "\r\n" + MessageBoxButtons.OK);

                Form2 loginw = new Form2();
                while (loginw.ShowDialog() != DialogResult.OK) ;
            }
        }
    }
}
