using System;
using System.IO;
using SAPbobsCOM;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        string excelPath =
            @"D:\Navinesh\VS_Application\RemoveVendorMasterREVA\ExcelRemoveVendor\DeleteBPReva.xlsx";

        string logPath =
            @"D:\Navinesh\VS_Application\RemoveVendorMasterREVA\ExcelRemoveVendor\Result.txt";

        Company company = new Company();

        company.Server = "SAPSERVER-1";
        company.CompanyDB = "REVA_LIVE";
        company.DbUserName = "sa";
        company.DbPassword = "Welcome1#";
        company.UserName = "manager";      // SAP User
        company.Password = "1234";      // SAP Password

        company.DbServerType = BoDataServerTypes.dst_MSSQL2019;
        company.language = BoSuppLangs.ln_English;

        int ret = company.Connect();     

        if (ret != 0)
        {
            company.GetLastError(out int errCode, out string errMsg);
            Console.WriteLine(errMsg);
            Console.ReadLine();
            return;
        }

        StreamWriter writer = new StreamWriter(logPath, false);

        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = excelApp.Workbooks.Open(excelPath);
        Excel.Worksheet sheet = workbook.Sheets["Sheet1"];

        int row = 2;

        while (true)
        {
            var cell = sheet.Cells[row, 1];

            if (cell.Value == null)
                break;

            string cardCode = cell.Value.ToString().Trim();

            try
            {
                BusinessPartners bp =
                    (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (bp.GetByKey(cardCode))
                {
                    if (bp.CardType != BoCardTypes.cSupplier)
                    {
                        writer.WriteLine(
                            $"{cardCode} : Not a Supplier");
                    }
                    else
                    {
                        int removeResult = bp.Remove();

                        if (removeResult == 0)
                        {
                            writer.WriteLine(
                                $"{cardCode} : Deleted Successfully");
                        }
                        else
                        {
                            company.GetLastError(
                                out int errorCode,
                                out string errorMessage);

                            writer.WriteLine(
                                $"{cardCode} : Failed - {errorMessage}");
                        }
                    }
                }
                else
                {
                    writer.WriteLine(
                        $"{cardCode} : Not Found");
                }
            }
            catch (Exception ex)
            {
                writer.WriteLine(
                    $"{cardCode} : Error - {ex.Message}");
            }

            row++;
        }

        writer.Close();

        workbook.Close(false);
        excelApp.Quit();

        company.Disconnect();

        Console.WriteLine("Process Completed.");
        Console.WriteLine("Result file: " + logPath);

        Console.ReadLine();
    }
}
