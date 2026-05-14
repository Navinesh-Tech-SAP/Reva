using System;
using System.Data.SqlClient;
using System.IO;

class Program
{
    static void Main()
    {
        string connectionString =
            "Server=SAPSERVER-1;Database=REVA_LIVE;User Id=sa;Password=Welcome1#;TrustServerCertificate=True;";

        string logFolder = @"D:\Navinesh\Log";
        string logFile = Path.Combine(logFolder, "FixedAssetDeleteLog.csv");

        try
        {
            if (!Directory.Exists(logFolder))
                Directory.CreateDirectory(logFolder);

            using (StreamWriter log = new StreamWriter(logFile, false))
            {
                log.WriteLine("No,ItemCode,Status");

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    Console.WriteLine("SQL Connection Successful ✅");

                    string selectQuery = "SELECT ItemCode FROM OITM WHERE ItemType = 'F'";

                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, con))
                    using (SqlDataReader dr = selectCmd.ExecuteReader())
                    {
                        int no = 1;

                        while (dr.Read())
                        {
                            string itemCode = dr["ItemCode"].ToString();

                            try
                            {
                                using (SqlConnection deleteCon = new SqlConnection(connectionString))
                                {
                                    deleteCon.Open();

                                    string deleteQuery = "DELETE FROM OITM WHERE ItemCode = @ItemCode AND ItemType = 'F'";

                                    using (SqlCommand deleteCmd = new SqlCommand(deleteQuery, deleteCon))
                                    {
                                        deleteCmd.Parameters.AddWithValue("@ItemCode", itemCode);

                                        int rows = deleteCmd.ExecuteNonQuery();

                                        if (rows > 0)
                                        {
                                            log.WriteLine($"{no},{itemCode},Deleted");
                                            Console.WriteLine($"{no} - {itemCode} - Deleted ✅");
                                        }
                                        else
                                        {
                                            log.WriteLine($"{no},{itemCode},Not Deleted");
                                            Console.WriteLine($"{no} - {itemCode} - Not Deleted ❌");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                log.WriteLine($"{no},{itemCode},Failed: {ex.Message.Replace(",", " ")}");
                                Console.WriteLine($"{no} - {itemCode} - Failed ❌ {ex.Message}");
                            }

                            no++;
                        }
                    }
                }
            }

            Console.WriteLine("Log generated at:");
            Console.WriteLine(logFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("MAIN ERROR ❌");
            Console.WriteLine(ex.Message);
        }

        Console.ReadLine();
    }
}
