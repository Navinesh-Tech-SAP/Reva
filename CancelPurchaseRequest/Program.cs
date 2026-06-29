using System;
using System.IO;
using SAPbobsCOM;

class Program
{
    static void Main(string[] args)
    {
        Company company = new Company();

        company.Server = "SAPSERVER-1";
        company.CompanyDB = "REVA_LIVE";
        company.DbUserName = "sa";
        company.DbPassword = "Welcome1#";

        company.UserName = "manager";
        company.Password = "1234";

        company.DbServerType = BoDataServerTypes.dst_MSSQL2019;
        company.language = BoSuppLangs.ln_English;

        int result = company.Connect();

        if (result != 0)
        {
            company.GetLastError(out int errCode, out string errMsg);
            Console.WriteLine($"Connection Failed: {errCode} - {errMsg}");
            Console.ReadLine();
            return;
        }

        Console.WriteLine("SAP Connected Successfully.");

        string logPath =
            @"D:\Navinesh\VS_Application\RemoveVendorMasterREVA\Logs\PRCancelLog.txt";

        Recordset rs =
            (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

        rs.DoQuery(
            "SELECT DocEntry, DocNum, CANCELED FROM OPRQ");

        using (StreamWriter sw = new StreamWriter(logPath, true))
        {
            sw.WriteLine("--------------------------------------");
            sw.WriteLine("Run Date : " + DateTime.Now);
            sw.WriteLine("--------------------------------------");

            while (!rs.EoF)
            {
                int docEntry =
                    Convert.ToInt32(rs.Fields.Item("DocEntry").Value);

                int docNum =
                    Convert.ToInt32(rs.Fields.Item("DocNum").Value);

                string canceled =
                    rs.Fields.Item("CANCELED").Value.ToString();

                sw.WriteLine(
                    $"DocNum : {docNum}  DocEntry : {docEntry}  Status : {canceled}");

                if (canceled == "N")
                {
                    Documents pr =
                        (Documents)company.GetBusinessObject(
                            (BoObjectTypes)1470000113);

                    if (pr.GetByKey(docEntry))
                    {
                        int ret = pr.Cancel();

                        if (ret == 0)
                        {
                            sw.WriteLine(
                                $"SUCCESS : PR {docNum} cancelled.");
                        }
                        else
                        {
                            company.GetLastError(
                                out int errCode,
                                out string errMsg);

                            sw.WriteLine(
                                $"FAILED : PR {docNum} - {errCode} - {errMsg}");
                        }
                    }
                }
                else
                {
                    sw.WriteLine(
                        $"Already Cancelled : PR {docNum}");
                }

                sw.WriteLine();

                rs.MoveNext();
            }
        }

        company.Disconnect();

        Console.WriteLine("Process Completed.");
        Console.WriteLine("Check log file:");
        Console.WriteLine(logPath);

        Console.ReadLine();
    }
}
