using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using SAPbobsCOM;

class Program
{
    static void Main(string[] args)
    {
        string excelPath = @"D:\Navinesh\DeleteBPMasterList\Cancel.xlsx";
        string logPath = @"D:\Navinesh\Log\result.txt";

        string connectionString = "Server=SAPSERVER-1;Database=LIVE_REVA_University;User Id=sa;Password=Welcome1#;TrustServerCertificate=True;";

        List<string> cardCodes = ReadExcel(excelPath);

        // SAP Company Connection (Assuming already configured)
        Company oCompany = ConnectToSAP();

        using (StreamWriter log = new StreamWriter(logPath, true))
        {
            foreach (string cardCode in cardCodes)
            {
                try
                {
                    if (!ExistsInSQL(connectionString, cardCode))
                    {
                        string msg = $" {cardCode} - Not Available in DB";
                        Console.WriteLine(msg);
                        log.WriteLine(msg);
                        continue;
                    }

                    BusinessPartners oBP = (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    if (oBP.GetByKey(cardCode))
                    {
                        int ret = oBP.Remove();

                        if (ret != 0)
                        {
                            oCompany.GetLastError(out int errCode, out string errMsg);

                            // fallback → Freeze BP
                            oBP.Frozen = BoYesNoEnum.tYES;
                            int upd = oBP.Update();

                            if (upd == 0)
                            {
                                string msg = $"⚠️ {cardCode} - Cannot delete, Frozen instead";
                                Console.WriteLine(msg);
                                log.WriteLine(msg);
                            }
                            else
                            {
                                string msg = $" {cardCode} - Error: {errMsg}";
                                Console.WriteLine(msg);
                                log.WriteLine(msg);
                            }
                        }
                        else
                        {
                            string msg = $" {cardCode} - Deleted";
                            Console.WriteLine(msg);
                            log.WriteLine(msg);
                        }
                    }
                    else
                    {
                        string msg = $" {cardCode} - Not found in SAP";
                        Console.WriteLine(msg);
                        log.WriteLine(msg);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
                }
                catch (Exception ex)
                {
                    string msg = $" {cardCode} - Exception: {ex.Message}";
                    Console.WriteLine(msg);
                    log.WriteLine(msg);
                }
            }
        }

        Console.WriteLine("Process Completed!");
        Console.ReadLine();
    }

    static List<string> ReadExcel(string path)
    {
        List<string> list = new List<string>();

        ExcelPackage.License.SetNonCommercialPersonal("Navinesh");

        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var ws = package.Workbook.Worksheets[0];
            int rowCount = ws.Dimension.Rows;

            for (int i = 2; i <= rowCount; i++)
            {
                string cardCode = ws.Cells[i, 1].Text.Trim();
                if (!string.IsNullOrEmpty(cardCode))
                    list.Add(cardCode);
            }
        }

        return list;
    }

    static bool ExistsInSQL(string connStr, string cardCode)
    {
        using (SqlConnection conn = new SqlConnection(connStr))
        {
            conn.Open();
            string query = "SELECT COUNT(*) FROM OCRD WHERE CardCode = @CardCode";

            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@CardCode", cardCode);
                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }
    }

    static Company ConnectToSAP()
    {
        Company oCompany = new Company
        {
            Server = "SAPSERVER-1",
            DbServerType = BoDataServerTypes.dst_MSSQL2019,
            CompanyDB = "LIVE_REVA_University",
            UserName = "manager",
            Password = "Cbs@2023",
            DbUserName = "sa",
            DbPassword = "Welcome1#",
            language = BoSuppLangs.ln_English
        };

        int ret = oCompany.Connect();

        if (ret != 0)
        {
            oCompany.GetLastError(out int errCode, out string errMsg);
            throw new Exception($"SAP Connection Failed: {errMsg}");
        }

        Console.WriteLine(" SAP Connected!");
        return oCompany;
    }
}
