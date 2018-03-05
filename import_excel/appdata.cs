using System.Data.SqlClient;

/// <summary>
/// Summary description for Class1
/// </summary>
/// 
namespace appdata
{
    public class appdatas
    {
        public appdatas()
        {
            //
            // TODO: Add constructor logic here
            connection = null;
            datatable = null;
            path = null;
            //
        }
        //member variables
        public SqlConnection connection;
        public string path;
        public System.Data.DataTable datatable;
        public System.Data.DataTable dtt;
        public SqlConnectionStringBuilder builder;
    }
}
