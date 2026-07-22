using SAPbobsCOM;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Xml.Linq;

class Program
{
    static Company oCompany;
    static string logFolder = @"D:\Navinesh\VS_Application\AttachmentInvoice_Transfer_Reva\Log";
    static string logFile;

    static void Main(string[] args)
    {
        ConnectSAP();

        // Prepare log file (one per run, with timestamp)
        if (!Directory.Exists(logFolder))
            Directory.CreateDirectory(logFolder);

        logFile = Path.Combine(logFolder, "AttachmentInvoiceLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");

        // Write header
        File.AppendAllText(logFile, "CardCode,DocNum,DocEntry,Status,Reason,Timestamp" + Environment.NewLine);

        // NOTE: Same PO-to-PO style logic, but for A/P Invoices (OPCH).
        // Attachment is pulled via T0.AtcEntry (the SOURCE invoice's own attachment entry) --
        // this is an Invoice-to-Invoice transfer, so the attachment must belong to the source
        // invoice itself. Source and target invoices are matched via T0.DocNum = T1.U_BaseNum
        // (per your query), NOT DocNum = DocNum like the PO version.
        // T1.DocEntry (target company DocEntry) is the key the DI API needs for GetByKey().
        string sql = @"
SELECT 
    T0.CardCode AS [CardCode],
    T0.DocNum AS [DocNum],
    T1.DocEntry AS [DocEntry],
    '\\172.21.0.45\SAP-Attachment\' AS [Path],
    AT.FileName + '.' + AT.FileExt AS [File Name]
FROM [LIVE_REVA_University].dbo.OPCH T0
INNER JOIN [LIVE_REVA_University].dbo.OATC OA
    ON T0.AtcEntry = OA.AbsEntry
INNER JOIN [LIVE_REVA_University].dbo.ATC1 AT
    ON OA.AbsEntry = AT.AbsEntry
INNER JOIN [REVA_LIVE].dbo.OPCH T1
    ON T0.DocNum = T1.U_BaseNum
WHERE
    T0.CANCELED = 'N'
    AND T1.CANCELED = 'N'
ORDER BY T0.CardCode;";

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
                string docNum = dr["DocNum"].ToString();
                int docEntry = Convert.ToInt32(dr["DocEntry"]);
                string sourcePath = dr["Path"].ToString().TrimEnd('\\'); // avoid double backslash
                string fullFileName = dr["File Name"].ToString();

                // split combined "name.ext" back into separate parts SAP needs
                string fileName = Path.GetFileNameWithoutExtension(fullFileName);
                string fileExt = Path.GetExtension(fullFileName).TrimStart('.');

                Console.WriteLine("----------------------------------------------------");
                Console.WriteLine("Processing -> CardCode: " + cardCode +
                                   " | DocNum: " + docNum +
                                   " | DocEntry: " + docEntry +
                                   " | File: " + fullFileName);

                bool success;
                string reason;

                AttachToInvoice(docEntry, fileName, fileExt, sourcePath, out success, out reason);

                if (success)
                    successCount++;
                else
                    failCount++;

                Console.WriteLine("Result -> " + (success ? "SUCCESS" : "FAILED") +
                                   (success ? "" : " | Reason: " + reason));

                WriteLog(cardCode, docNum, docEntry, success ? "Success" : "Failed", reason);
            }
            dr.Close();
        }

        Console.WriteLine("Completed. Total: " + totalCount + " | Success: " + successCount + " | Failed: " + failCount);
        Console.WriteLine("Log file created at : " + logFile);
        Console.ReadLine();
    }

    static void AttachToInvoice(int docEntry,
                                 string fileName,
                                 string fileExt,
                                 string sourcePath,
                                 out bool success,
                                 out string reason)
    {
        Documents oInv = null;
        Attachments2 oATT = null;
        success = false;
        reason = "";

        try
        {
            oInv = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);

            if (!oInv.GetByKey(docEntry))
            {
                reason = "Invoice Not Found";
                Console.WriteLine("Invoice Not Found : DocEntry " + docEntry);
                return;
            }

            oATT = (Attachments2)oCompany.GetBusinessObject(BoObjectTypes.oAttachments2);

            int attachEntry = oInv.AttachmentEntry;

            if (attachEntry > 0 && oATT.GetByKey(attachEntry))
            {
                // Check if this exact file (name + extension) is already attached
                // to avoid creating duplicate lines on repeat runs.
                bool alreadyExists = false;
                int lineCount = oATT.Lines.Count;

                for (int i = 0; i < lineCount; i++)
                {
                    oATT.Lines.SetCurrentLine(i);
                    if (string.Equals(oATT.Lines.FileName, fileName, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(oATT.Lines.FileExtension, fileExt, StringComparison.OrdinalIgnoreCase))
                    {
                        alreadyExists = true;
                        break;
                    }
                }

                if (alreadyExists)
                {
                    reason = "Already Attached";
                    success = true; // not a failure, just nothing to do
                    Console.WriteLine("Skipped (already attached) : " + fileName + "." + fileExt);
                    return;
                }

                // Invoice already has an attachment record -> add a new line to it
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
                Console.WriteLine("Attachment updated successfully for Invoice DocEntry : " + docEntry);
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
                oInv.AttachmentEntry = newAttachEntry;

                int invUpdateResult = oInv.Update();

                if (invUpdateResult != 0)
                {
                    reason = oCompany.GetLastErrorDescription();
                    Console.WriteLine("Invoice Update Failed : " + reason);
                    return;
                }

                success = true;
                Console.WriteLine("Attachment added successfully for Invoice DocEntry : " + docEntry);
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
            if (oInv != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oInv);
                oInv = null;
            }
            GC.Collect();
        }
    }

    static void WriteLog(string cardCode, string docNum, int docEntry, string status, string reason)
    {
        try
        {
            // escape commas/quotes so CSV doesn't break if reason contains them
            string safeReason = reason?.Replace("\"", "'").Replace(",", ";") ?? "";

            string line = string.Format("{0},{1},{2},{3},{4},{5}",
                cardCode,
                docNum,
                docEntry,
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
