
using System;
using System.Data.SqlClient;

class Program
{
    static void Main()
    {
        string connectionString =
            "Server=SAPSERVER-1;" +
            "Database=RevaIncomingPayment;" +
            "User Id=sa;" +
            "Password=Welcome1#;" +
            "TrustServerCertificate=True;";

        try
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                Console.WriteLine("SQL Connection Successful!");

                SqlCommand cmd = new SqlCommand("SELECT @@SERVERNAME", conn);
                string serverName = Convert.ToString(cmd.ExecuteScalar());

                Console.WriteLine("Connected Server: " + serverName);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Connection Failed!");
            Console.WriteLine(ex.Message);
        }

        Console.ReadLine();
    }
}
