using System;
using System.Data.SqlClient;
using System.IO;
using SAPbobsCOM;

class Program
{
    static Company oCompany;
    static string logFolder = @"D:\Navinesh\VS_Application\AttachmentBP_Transfer_Reva\Log";
    static string logFile;

    static void Main(string[] args)
    {
        ConnectSAP();

        // Prepare log file (one per run, with timestamp)
        if (!Directory.Exists(logFolder))
            Directory.CreateDirectory(logFolder);

        logFile = Path.Combine(logFolder, "AttachmentLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");

        // Write header
        File.AppendAllText(logFile, "CardCode,Status,Reason,Timestamp" + Environment.NewLine);

        string sql = @"
SELECT 
   T0.""CardCode"",
   '\\172.21.0.45\SAP-Attachment\' AS ""Path"",
   T2.""FileName"" + '.' + T2.FileExt AS ""FullFileName""
FROM [LIVE_REVA_University].dbo.OCRD T0
INNER JOIN [LIVE_REVA_University].dbo.OATC T1
    ON T0.""AtcEntry"" = T1.""AbsEntry""
INNER JOIN [LIVE_REVA_University].dbo.ATC1 T2
    ON T1.""AbsEntry"" = T2.""AbsEntry""
INNER JOIN [REVA_LIVE].dbo.OCRD T3
    ON T3.CardCode = T0.CardCode AND T3.""CardName"" = T0.""CardName""
ORDER BY T0.""CardCode"";";

        string conn =
            @"Server=SAPSERVER-1;
              Database=LIVE_REVA_University;
              User ID=sa;
              Password=Welcome1#;
              TrustServerCertificate=True;";

        int totalCount = 0;
        int successCount = 0;
        int failCount = 0;

        using (SqlConnection cn = new SqlConnection(conn))
        {
            cn.Open();
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                totalCount++;

                string cardCode = dr["CardCode"].ToString();
                string sourcePath = dr["Path"].ToString().TrimEnd('\\'); // avoid double backslash
                string fullFileName = dr["FullFileName"].ToString();

                // split combined "name.ext" back into separate parts SAP needs
                string fileName = Path.GetFileNameWithoutExtension(fullFileName);
                string fileExt = Path.GetExtension(fullFileName).TrimStart('.');

                bool success;
                string reason;

                AttachToBP(cardCode, fileName, fileExt, sourcePath, out success, out reason);

                if (success)
                    successCount++;
                else
                    failCount++;

                WriteLog(cardCode, success ? "Success" : "Failed", reason);
            }
            dr.Close();
        }

        Console.WriteLine("Completed. Total: " + totalCount + " | Success: " + successCount + " | Failed: " + failCount);
        Console.WriteLine("Log file created at : " + logFile);
        Console.ReadLine();
    }

    static void AttachToBP(string cardCode,
                            string fileName,
                            string fileExt,
                            string sourcePath,
                            out bool success,
                            out string reason)
    {
        BusinessPartners bp = null;
        Attachments2 oATT = null;
        success = false;
        reason = "";

        try
        {
            bp = (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            if (!bp.GetByKey(cardCode))
            {
                reason = "BP Not Found";
                Console.WriteLine("BP Not Found : " + cardCode);
                return;
            }

            oATT = (Attachments2)oCompany.GetBusinessObject(BoObjectTypes.oAttachments2);

            int attachEntry = bp.AttachmentEntry;

            if (attachEntry > 0 && oATT.GetByKey(attachEntry))
            {
                // BP already has an attachment record -> add a new line to it
                oATT.Lines.Add();
                oATT.Lines.FileName = fileName;
                oATT.Lines.FileExtension = fileExt;
                oATT.Lines.SourcePath = sourcePath;
                oATT.Lines.Override = BoYesNoEnum.tYES;

                int updateResult = oATT.Update();

                if (updateResult != 0)
                {
                    reason = oCompany.GetLastErrorDescription();
                    Console.WriteLine("Attachment Update Failed : " + reason);
                    return;
                }

                success = true;
                Console.WriteLine("Attachment updated successfully for BP : " + cardCode);
            }
            else
            {
                // No existing attachment -> create a new one
                oATT.Lines.Add();
                oATT.Lines.FileName = fileName;
                oATT.Lines.FileExtension = fileExt;
                oATT.Lines.SourcePath = sourcePath;
                oATT.Lines.Override = BoYesNoEnum.tYES;
                oATT.Lines.CopyToTargetDoc = BoYesNoEnum.tYES;

                int addResult = oATT.Add();

                if (addResult != 0)
                {
                    reason = oCompany.GetLastErrorDescription();
                    Console.WriteLine("Attachment Add Failed : " + reason);
                    return;
                }

                int newAttachEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                bp.AttachmentEntry = newAttachEntry;

                int bpUpdateResult = bp.Update();

                if (bpUpdateResult != 0)
                {
                    reason = oCompany.GetLastErrorDescription();
                    Console.WriteLine("BP Update Failed : " + reason);
                    return;
                }

                success = true;
                Console.WriteLine("Attachment added successfully for BP : " + cardCode);
            }
        }
        catch (Exception ex)
        {
            reason = "Exception: " + ex.Message;
            Console.WriteLine("Attachment Error : " + ex.Message);
        }
        finally
        {
            if (oATT != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oATT);
                oATT = null;
            }
            if (bp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(bp);
                bp = null;
            }
            GC.Collect();
        }
    }

    static void WriteLog(string cardCode, string status, string reason)
    {
        try
        {
            // escape commas/quotes so CSV doesn't break if reason contains them
            string safeReason = reason?.Replace("\"", "'").Replace(",", ";") ?? "";

            string line = string.Format("{0},{1},{2},{3}",
                cardCode,
                status,
                safeReason,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            File.AppendAllText(logFile, line + Environment.NewLine);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Log Write Failed : " + ex.Message);
        }
    }

    static void ConnectSAP()
    {
        oCompany = new Company();
        oCompany.Server = "SAPSERVER-1";
        oCompany.CompanyDB = "REVA_LIVE";
        oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019;
        oCompany.DbUserName = "sa";
        oCompany.DbPassword = "Welcome1#";
        oCompany.UserName = "manager";
        oCompany.Password = "1234";
        oCompany.LicenseServer = "SAPSERVER-1:30000";

        int ret = oCompany.Connect();
        if (ret != 0)
        {
            Console.WriteLine(oCompany.GetLastErrorDescription());
            Environment.Exit(0);
        }
        Console.WriteLine("Connected Successfully.");
    }
}
