
using System;
using Microsoft.Data.SqlClient;

class Program
{
    static void Main()
    {
        //string connectionString = "Data Source=SAPSERVER-1;Initial Catalog=REVA_PORTAL_LIVE;Integrated Security=True";
        //OR use SQL login:
        string connectionString = "Server=SAPSERVER-1;Database=REVA_PORTAL_LIVE;User Id=sa;Password=Welcome1#;TrustServerCertificate=True;";

        try
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                Console.WriteLine("SQL Server Connected Successfully ✅");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Connection Failed ❌");
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
