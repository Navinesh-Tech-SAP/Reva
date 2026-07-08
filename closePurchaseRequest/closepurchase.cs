using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace CloseOPRQ_Reva
{
    class OprqRow
    {
        public int RowIndex { get; set; }
        public int DocNum { get; set; }
        public string Status { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Company company = new Company();

            try
            {
                // SAP Connection
                company.Server = ConfigurationManager.AppSettings["Server"];
                company.CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
                company.DbUserName = ConfigurationManager.AppSettings["DbUserName"];
                company.DbPassword = ConfigurationManager.AppSettings["DbPassword"];
                company.UserName = ConfigurationManager.AppSettings["SapUserName"];
                company.Password = ConfigurationManager.AppSettings["SapPassword"];
                company.LicenseServer = ConfigurationManager.AppSettings["LicenseServer"];

                company.DbServerType = BoDataServerTypes.dst_MSSQL2019;
                company.language = BoSuppLangs.ln_English;

                int connectResult = company.Connect();

                if (connectResult != 0)
                {
                    company.GetLastError(out int errorCode, out string errorMessage);

                    Console.WriteLine("SAP Connection Failed");
                    Console.WriteLine(errorCode);
                    Console.WriteLine(errorMessage);
                    return;
                }

                Console.WriteLine("Connected Successfully");
                Console.WriteLine("Company : " + company.CompanyName);


                string excelPath =
                    @"D:\Navinesh\VS_Application\CloseGRPO_Reva\Excel\CancelOPRQ.xlsx";


                List<OprqRow> oprqRows = ReadExcelColumnA(excelPath);

                if (oprqRows == null || oprqRows.Count == 0)
                {
                    Console.WriteLine("Invalid Purchase Request number in Excel.");
                    return;
                }

                foreach (OprqRow row in oprqRows)
                {
                    int docNum = row.DocNum;

                    if (docNum == 0)
                    {
                        row.Status = "Invalid DocNum";
                        Console.WriteLine(docNum + " - " + row.Status);
                        continue;
                    }

                    // Get DocEntry from DocNum
                    Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    rs.DoQuery($@"
SELECT DocEntry
FROM OPRQ
WHERE DocNum = {docNum}");

                    if (rs.EoF)
                    {
                        row.Status = "Not Found";
                        Console.WriteLine(docNum + " - " + row.Status);
                        continue;
                    }

                    int docEntry = Convert.ToInt32(
                        rs.Fields.Item("DocEntry").Value
                    );


                    Documents oprq =
                        (Documents)company.GetBusinessObject(
                            BoObjectTypes.oPurchaseRequest);


                    if (!oprq.GetByKey(docEntry))
                    {
                        company.GetLastError(
                            out int errorCode,
                            out string errorMessage);

                        row.Status = "Load Failed";
                        Console.WriteLine(docNum + " - " + row.Status);
                        continue;
                    }


                    // Check if the Purchase Request is already Closed before attempting closure
                    if (oprq.DocumentStatus == BoStatus.bost_Close)
                    {
                        row.Status = "Already Closed";
                        Console.WriteLine(docNum + " - " + row.Status);
                        continue;
                    }


                    // Close Purchase Request
                    int closeResult = oprq.Close();

                    if (closeResult == 0)
                    {
                        row.Status = "Closed";
                    }
                    else
                    {
                        company.GetLastError(
                            out int errorCode,
                            out string errorMessage);

                        row.Status = "Open";
                    }

                    Console.WriteLine(docNum + " - " + row.Status);
                }

                // Write DocNum/Status log back into the same Excel file (Column B)
                WriteStatusToExcel(excelPath, oprqRows);

                // Export a DocNum,Status log file into the LOG folder
                ExportLogFile(oprqRows);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception : " + ex.Message);
            }
            finally
            {
                if (company.Connected)
                {
                    company.Disconnect();
                    Console.WriteLine("Disconnected");
                }
            }

            Console.ReadLine();
        }



        static List<OprqRow> ReadExcelColumnA(string filePath)
        {
            List<OprqRow> docNums = new List<OprqRow>();

            try
            {
                using (SpreadsheetDocument document =
                    SpreadsheetDocument.Open(filePath, false))
                {

                    WorkbookPart workbookPart =
                        document.WorkbookPart;


                    Sheet sheet =
                        workbookPart.Workbook.Sheets
                        .Elements<Sheet>()
                        .FirstOrDefault(x => x.Name == "Sheet1");


                    if (sheet == null)
                    {
                        Console.WriteLine("Sheet1 not found");
                        return docNums;
                    }


                    WorksheetPart worksheetPart =
                        (WorksheetPart)workbookPart
                        .GetPartById(sheet.Id);


                    // Get all rows in the sheet, ordered by row index
                    var rows = worksheetPart.Worksheet
                        .Descendants<Row>()
                        .OrderBy(r => (int)r.RowIndex.Value)
                        .ToList();


                    foreach (Row row in rows)
                    {
                        Cell cell = row.Elements<Cell>()
                            .FirstOrDefault(c =>
                                GetColumnName(c.CellReference) == "A");


                        if (cell == null)
                        {
                            // Empty cell in column A for this row -> stop
                            break;
                        }


                        string value =
                            GetCellValue(document, cell);


                        if (string.IsNullOrWhiteSpace(value))
                        {
                            // Blank cell -> stop reading
                            break;
                        }


                        if (int.TryParse(value, out int docNum))
                        {
                            docNums.Add(new OprqRow
                            {
                                RowIndex = (int)row.RowIndex.Value,
                                DocNum = docNum
                            });
                        }
                        else
                        {
                            // Non-numeric value -> stop reading
                            break;
                        }
                    }


                    return docNums;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel Error : " + ex.Message);
                return docNums;
            }
        }



        static void WriteStatusToExcel(string filePath, List<OprqRow> oprqRows)
        {
            try
            {
                using (SpreadsheetDocument document =
                    SpreadsheetDocument.Open(filePath, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;

                    Sheet sheet =
                        workbookPart.Workbook.Sheets
                        .Elements<Sheet>()
                        .FirstOrDefault(x => x.Name == "Sheet1");

                    if (sheet == null)
                    {
                        Console.WriteLine("Sheet1 not found. Could not write status.");
                        return;
                    }

                    WorksheetPart worksheetPart =
                        (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                    SheetData sheetData =
                        worksheetPart.Worksheet.Elements<SheetData>().First();

                    foreach (OprqRow oprqRow in oprqRows)
                    {
                        uint rowIndex = (uint)oprqRow.RowIndex;

                        // Find or create the Row
                        Row excelRow = sheetData.Elements<Row>()
                            .FirstOrDefault(r => r.RowIndex.Value == rowIndex);

                        if (excelRow == null)
                        {
                            excelRow = new Row() { RowIndex = rowIndex };
                            sheetData.Append(excelRow);
                        }

                        string cellReference = "B" + rowIndex;

                        // Find or create the Cell in column B
                        Cell statusCell = excelRow.Elements<Cell>()
                            .FirstOrDefault(c => c.CellReference == cellReference);

                        if (statusCell == null)
                        {
                            statusCell = new Cell()
                            {
                                CellReference = cellReference
                            };

                            // Insert in correct column order
                            Cell nextCell = excelRow.Elements<Cell>()
                                .FirstOrDefault(c =>
                                    string.Compare(
                                        GetColumnName(c.CellReference),
                                        "B",
                                        StringComparison.OrdinalIgnoreCase) > 0);

                            if (nextCell != null)
                            {
                                excelRow.InsertBefore(statusCell, nextCell);
                            }
                            else
                            {
                                excelRow.Append(statusCell);
                            }
                        }

                        statusCell.DataType = CellValues.String;
                        statusCell.CellValue =
                            new CellValue(oprqRow.Status ?? string.Empty);
                    }

                    // Optional header for column B
                    Row headerRow = sheetData.Elements<Row>()
                        .FirstOrDefault(r => r.RowIndex.Value == 1);

                    if (headerRow != null)
                    {
                        Cell headerCell = headerRow.Elements<Cell>()
                            .FirstOrDefault(c => c.CellReference == "B1");

                        if (headerCell == null)
                        {
                            headerCell = new Cell() { CellReference = "B1" };
                            headerRow.Append(headerCell);
                        }

                        // Only set header text if B1 doesn't already contain the DocNum data
                        // (i.e. treat row 1 as a header row for the log column)
                        headerCell.DataType = CellValues.String;
                        headerCell.CellValue = new CellValue("Status");
                    }

                    worksheetPart.Worksheet.Save();
                }

                Console.WriteLine("Status log written back to Excel : " + filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel Write Error : " + ex.Message);
            }
        }



        static void ExportLogFile(List<OprqRow> oprqRows)
        {
            try
            {
                string logFolder =
                    @"D:\Navinesh\VS_Application\CloseGRPO_Reva\LOG";

                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }

                string fileName =
                    "OPRQ_CloseLog_" +
                    DateTime.Now.ToString("yyyyMMdd_HHmmss") +
                    ".csv";

                string logFilePath = Path.Combine(logFolder, fileName);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("DocNum,Status");

                foreach (OprqRow row in oprqRows)
                {
                    sb.AppendLine(row.DocNum + "," + (row.Status ?? "Unknown"));
                }

                File.WriteAllText(logFilePath, sb.ToString());

                Console.WriteLine("Log exported to : " + logFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log Export Error : " + ex.Message);
            }
        }



        static string GetColumnName(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return string.Empty;

            return new string(
                cellReference.TakeWhile(c => !char.IsDigit(c)).ToArray());
        }



        static string GetCellValue(
            SpreadsheetDocument document,
            Cell cell)
        {
            string value = cell.InnerText;


            if (cell.DataType == null)
                return value;


            if (cell.DataType.Value ==
                CellValues.SharedString)
            {
                return document.WorkbookPart
                    .SharedStringTablePart
                    .SharedStringTable
                    .Elements<SharedStringItem>()
                    .ElementAt(int.Parse(value))
                    .InnerText;
            }


            return value;
        }
    }
}
