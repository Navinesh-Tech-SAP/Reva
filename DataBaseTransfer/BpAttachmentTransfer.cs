// Requires NuGet package: ClosedXML
// Install-Package ClosedXML
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using SAPbobsCOM;

class Program
{
    static SAPbobsCOM.Company oCompany;

    static string excelInputFile = @"D:\Navinesh\VS_Application\AttachmentTransferREVA\Excel\BPAttachment.xlsx";
    static string logFolder = @"D:\Navinesh\VS_Application\AttachmentTransferREVA\Log";
    static string logFile;

    // Result row collected for each BP, written out to the Excel log at the end
    class BpResult
    {
        public string CardCode;
        public string Status;   // Success / Failed
        public string Reason;   // blank if success
    }

    static void Main(string[] args)
    {
        ConnectSAP();

        if (!Directory.Exists(logFolder))
            Directory.CreateDirectory(logFolder);

        logFile = Path.Combine(logFolder, "BPAttachmentLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

        List<string> cardCodes = ReadCardCodesFromExcel(excelInputFile);

        if (cardCodes.Count == 0)
        {
            Console.WriteLine("No BP codes found in Excel file : " + excelInputFile);
            Console.ReadLine();
            return;
        }

        Console.WriteLine("Loaded " + cardCodes.Count + " BP code(s) from Excel.");

        List<BpResult> results = new List<BpResult>();

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

            foreach (string cardCode in cardCodes)
            {
                totalCount++;
                Console.WriteLine("----------------------------------------------------");
                Console.WriteLine("Processing BP : " + cardCode);

                // Get all attachment files (from the source DB) linked to this BP's attachment record
                string sql = @"
SELECT 
    '\\172.21.0.45\SAP-Attachment\' AS [Path],
    AT.FileName + '.' + AT.FileExt AS [File Name]
FROM [LIVE_REVA_University].dbo.OCRD BP
INNER JOIN [LIVE_REVA_University].dbo.OATC OA
    ON BP.AtcEntry = OA.AbsEntry
INNER JOIN [LIVE_REVA_University].dbo.ATC1 AT
    ON OA.AbsEntry = AT.AbsEntry
WHERE BP.CardCode = @CardCode;";

                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.Parameters.AddWithValue("@CardCode", cardCode);

                List<(string sourcePath, string fullFileName)> files = new List<(string, string)>();

                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string sourcePath = dr["Path"].ToString().TrimEnd('\\');
                        string fullFileName = dr["File Name"].ToString();
                        files.Add((sourcePath, fullFileName));
                    }
                }

                if (files.Count == 0)
                {
                    Console.WriteLine("Result -> FAILED | Reason: No attachment found in source for this BP");
                    failCount++;
                    results.Add(new BpResult { CardCode = cardCode, Status = "Failed", Reason = "No attachment found in source for this BP" });
                    continue;
                }

                bool bpOverallSuccess = true;
                List<string> bpReasons = new List<string>();

                foreach (var file in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file.fullFileName);
                    string fileExt = Path.GetExtension(file.fullFileName).TrimStart('.');

                    bool success;
                    string reason;

                    AttachToBP(cardCode, fileName, fileExt, file.sourcePath, out success, out reason);

                    Console.WriteLine("  File: " + file.fullFileName +
                                       " -> " + (success ? "SUCCESS" : "FAILED") +
                                       (string.IsNullOrEmpty(reason) ? "" : " | " + reason));

                    if (!success)
                    {
                        bpOverallSuccess = false;
                        bpReasons.Add(file.fullFileName + ": " + reason);
                    }
                }

                if (bpOverallSuccess)
                {
                    successCount++;
                    Console.WriteLine("Result -> SUCCESS (all files transferred for " + cardCode + ")");
                    results.Add(new BpResult { CardCode = cardCode, Status = "Success", Reason = "" });
                }
                else
                {
                    failCount++;
                    string combinedReason = string.Join(" | ", bpReasons);
                    Console.WriteLine("Result -> FAILED | Reason: " + combinedReason);
                    results.Add(new BpResult { CardCode = cardCode, Status = "Failed", Reason = combinedReason });
                }
            }
        }

        WriteExcelLog(results);

        Console.WriteLine("====================================================");
        Console.WriteLine("Completed. Total BPs: " + totalCount + " | Success: " + successCount + " | Failed: " + failCount);
        Console.WriteLine("Excel log created at : " + logFile);
        Console.ReadLine();
    }

    // Reads BP codes from Sheet1, starting at cell A1, going down until an empty cell is hit
    static List<string> ReadCardCodesFromExcel(string filePath)
    {
        List<string> codes = new List<string>();

        if (!File.Exists(filePath))
        {
            Console.WriteLine("Excel input file not found : " + filePath);
            return codes;
        }

        using (var workbook = new XLWorkbook(filePath))
        {
            var ws = workbook.Worksheet(1); // Sheet1

            int row = 1;
            while (true)
            {
                var cell = ws.Cell(row, 1); // Column A
                string value = cell.GetString().Trim();

                if (string.IsNullOrEmpty(value))
                    break;

                codes.Add(value);
                row++;
            }
        }

        return codes;
    }

    // Writes BP Code, Status, Reason for every processed BP to a new Excel log file
    static void WriteExcelLog(List<BpResult> results)
    {
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Log");

                ws.Cell(1, 1).Value = "BP Code";
                ws.Cell(1, 2).Value = "Status";
                ws.Cell(1, 3).Value = "Reason";
                ws.Cell(1, 4).Value = "Timestamp";
                ws.Row(1).Style.Font.Bold = true;

                int row = 2;
                foreach (var r in results)
                {
                    ws.Cell(row, 1).Value = r.CardCode;
                    ws.Cell(row, 2).Value = r.Status;
                    ws.Cell(row, 3).Value = r.Reason;
                    ws.Cell(row, 4).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    row++;
                }

                ws.Columns().AdjustToContents();
                workbook.SaveAs(logFile);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Excel Log Write Failed : " + ex.Message);
        }
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
                return;
            }

            oATT = (Attachments2)oCompany.GetBusinessObject(BoObjectTypes.oAttachments2);

            int attachEntry = bp.AttachmentEntry;

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
                    success = true;
                    return;
                }

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
                    return;
                }

                success = true;
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
                    return;
                }

                int newAttachEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                bp.AttachmentEntry = newAttachEntry;

                int bpUpdateResult = bp.Update();

                if (bpUpdateResult != 0)
                {
                    reason = oCompany.GetLastErrorDescription();
                    return;
                }

                success = true;
            }
        }
        catch (Exception ex)
        {
            reason = "Exception: " + ex.Message;
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

    static void ConnectSAP()
    {
        oCompany = new SAPbobsCOM.Company();
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
