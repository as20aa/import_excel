using System;
using System.Data;
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
            //这段代码可以实现打开程序必须先进入登陆界面，然后才能进行数据导入这个功能
            data = new appdatas();
            Form2 loginw = new Form2();
            if (loginw.ShowDialog() != DialogResult.OK)
            {
                //当取消登陆时，整个程序将会退出
                this.Close();
            }
            if(data.connection.State==ConnectionState.Open)
                label2.Text = "已连接数据库";
        }

        //选择目标文件
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

        //将数据导入到datatable中，此处使用的是excel的com组件
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

            //在datagridview中显示数据，设置数据源之后界面会自动刷新
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

        //将数据从datatable导入到SQL Server上，并且对数据库、表格、记录进行检查
        private void itsql(object sender, EventArgs e)
        {
            //将连接的任务交给import
            import ipf = new import();
            ipf.Show();
            //try
            //{
            //    //连接数据库
            //    using (data.connection = new SqlConnection(data.builder.ConnectionString))
            //    {
            //        //登陆成功
            //        data.connection.Open();

            //        //打开导入界面
            //        import ipf= new import();
            //        ipf.Show();

            ////测试能够从与SQL Server的连接中获取数据库列表
            //StringBuilder getlist = new StringBuilder();
            //getlist.Append("SELECT Name FROM " + data.builder.InitialCatalog + "..SysDatabases ORDER BY Name");
            //DataTable databaselist = new DataTable();
            //SqlDataAdapter list = new SqlDataAdapter(getlist.ToString(), data.connection);
            //list.Fill(databaselist);
            //if (databaselist.Rows.Count > 0)
            //{
            //    MessageBox.Show("数据库非空");
            //}
            ////导入到SQL Server
            ////检查数据库
            //string sql = "if not exists (select * from sys.databases where name = 'list') create database [list];";
            //using (SqlCommand command = new SqlCommand(sql, data.connection))
            //{
            //    command.ExecuteNonQuery();
            //}
            //StringBuilder sb = new StringBuilder();
            //sb.Append("use list;");
            //sb.Append("if not exists (select * from sysobjects where name='datatable' and xtype='U') create table datatable (");

            ////添加列信息
            //foreach (DataColumn columns in data.datatable.Columns)
            //{
            //    sb.Append(" ");
            //    //插入的数据的列名称
            //    sb.Append(columns.ColumnName.ToString());
            //    sb.Append(" ");
            //    //插入列的数据类型
            //    if (columns.DataType == System.Type.GetType("System.String"))
            //        sb.Append("nvarchar(50)");
            //    if (columns.DataType == System.Type.GetType("System.Int64"))
            //        sb.Append("bigint");
            //    else
            //    {
            //        if (columns.DataType == System.Type.GetType("System.Int32"))
            //            sb.Append("int");
            //        else
            //        {
            //            if (columns.DataType == System.Type.GetType("System.Int16"))
            //                sb.Append("smallint");
            //            else
            //                if (columns.DataType == System.Type.GetType("System.Double"))
            //                sb.Append("real");
            //        }
            //    }
            //    if (columns.DataType == System.Type.GetType("System.DateTime"))
            //        sb.Append("datetime");
            //    sb.Append(",");
            //}
            ////移除最后一个','
            //sb = sb.Remove(sb.Length - 1, 1);
            //sb.Append(");");
            //sql = sb.ToString();
            ////创建SQL数据库操作
            //using (SqlCommand command = new SqlCommand(sql, data.connection))
            //{
            //    command.ExecuteNonQuery();
            //    //Console.WriteLine("Table creation is done.");
            //}

            //从datatable的每一行开始筛选，对记录进行查重
            //foreach (DataRow row2 in data.datatable.Rows)
            //{
            //    StringBuilder jd = new StringBuilder();
            //    //向SQL Server发出查询指令，并根据返回数据判断是否存在重复记录
            //    jd.Append("select * from datatable where ");
            //    jd.Append(data.datatable.Columns[0].ColumnName);
            //    jd.Append(" = ");
            //    jd.Append(row2[data.datatable.Columns[0].ColumnName]);
            //    jd.Append(";");
            //    sql = jd.ToString();
            //    DataSet dataset = new DataSet();
            //    //SqlDataAdapter类的用法，接收到的数据是DataSet或者是DataTable类型
            //    SqlDataAdapter adapter = new SqlDataAdapter(sql, data.connection);

            //    //将接收到的数据填充到dataset中
            //    adapter.Fill(dataset);
            //    //采用以下的方式才是正确的对datatable和dataset单元数据的读取方式
            //    //MessageBox.Show(temp.Rows[0][0].ToString());
            //    //MessageBox.Show(dataset.Tables[0].Rows[0][0].ToString());
            //    //adapter.SelectCommand = new SqlCommand(sql, data.connection);
            //    //使用dataset中表的行数判断查询结果是否为空
            //    if (dataset.Tables[0].Rows.Count > 0)
            //    {

            //        //一旦dataset为空，则下面的判断语句就会出错！
            //        if (row2[0].ToString() == dataset.Tables[0].Rows[0][0].ToString())
            //        {
            //            //弹出信息框，对信息框方法的描述见定义
            //            //注意添加的按钮属性是okcancel而不是yesno
            //            DialogResult dr = MessageBox.Show("表格重复！是否更新该列表？", "警告", MessageBoxButtons.OKCancel);
            //            if (dr == DialogResult.OK)
            //            {
            //                //更新该列表
            //                StringBuilder di = new StringBuilder();
            //                di.Append("delete from datatable where ");
            //                di.Append(data.datatable.Columns[0].ColumnName);
            //                di.Append(" = ");
            //                di.Append(row2[data.datatable.Columns[0]].ToString());
            //                di.Append(";");
            //                sql = di.ToString();
            //                using (SqlCommand command2 = new SqlCommand(sql, data.connection))
            //                {
            //                    command2.ExecuteNonQuery();
            //                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(data.connection))
            //                    {
            //                        bulkCopy.DestinationTableName = "dbo.datatable";
            //                        // Write from the source to the destination.
            //                        try
            //                        {
            //                            //从当前的datatable中克隆到新的datatable中，克隆得到的datatable具有原来
            //                            //datatable的架构，但是没有数据，即只有列，没有行数据
            //                            DataTable buff = data.datatable.Clone();
            //                            //不能用buff.rows.add(row2),程序会直接出错，可以中buff.importrow(row2)来实现导入一行
            //                            buff.ImportRow(row2);
            //                            // Write from the source to the destination.
            //                            bulkCopy.WriteToServer(buff);
            //                            //由于buff是局部变量，在每次调用之后会被自动清除故可以不用移除操作
            //                            //buff.Rows.Remove(row2);
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            Console.WriteLine(ex.Message);
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //    else
            //    {
            //        //当不存在重复行数据时
            //        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(data.connection))
            //        {
            //            bulkCopy.DestinationTableName = "dbo.datatable";
            //            try
            //            {
            //                DataTable buff = data.datatable.Clone();
            //                //不能用buff.rows.add(row2),程序会直接出错，可以中buff.importrow(row2)来实现导入一行
            //                buff.ImportRow(row2);
            //                // Write from the source to the destination.
            //                bulkCopy.WriteToServer(buff);
            //                //buff.Rows.Remove(row2);
            //            }
            //            catch (Exception ex)
            //            {
            //                Console.WriteLine(ex.Message);
            //            }
            //        }
            //    }

            //}
    //    }
    //}
    //        catch
    //        {
    //            label2.Text = "数据库未连接";
    //            MessageBox.Show("登陆失败！请重试" + "\r\n" + MessageBoxButtons.OK);

    //            Form2 loginw = new Form2();
    //            while (loginw.ShowDialog() != DialogResult.OK) ;
    //        }
        }
    }
}
