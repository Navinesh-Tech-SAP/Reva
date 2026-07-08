using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace CancelGRPO_Reva
{
    // Simple holder for a DocNum plus the Excel row it came from,
    // and the eventual status we log back to the sheet
    class GrpoRow
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
                company.language = BoSuppLangs.ln_English; ;

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
                    @"D:\Navinesh\VS_Application\CloseGRPO_Reva\Excel\CancelGRPO.xlsx";


                List<GrpoRow> grpoRows = ReadExcelColumnA(excelPath);

                if (grpoRows == null || grpoRows.Count == 0)
                {
                    Console.WriteLine("Invalid GRPO number in Excel.");
                    return;
                }

                foreach (GrpoRow row in grpoRows)
                {
                    int docNum = row.DocNum;

                    if (docNum == 0)
                    {
                        Console.WriteLine("Invalid GRPO number in Excel.");
                        row.Status = "Invalid DocNum";
                        continue;
                    }

                    Console.WriteLine("GRPO DocNum : " + docNum);


                    // Get DocEntry from DocNum
                    Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    rs.DoQuery($@"
SELECT DocEntry
FROM OPDN
WHERE DocNum = {docNum}");

                    if (rs.EoF)
                    {
                        Console.WriteLine("GRPO not found in OPDN");
                        row.Status = "Not Found";
                        continue;
                    }


                    int docEntry = Convert.ToInt32(
                        rs.Fields.Item("DocEntry").Value
                    );


                    Console.WriteLine("GRPO DocEntry : " + docEntry);



                    Documents grpo =
                        (Documents)company.GetBusinessObject(
                            BoObjectTypes.oPurchaseDeliveryNotes);


                    if (!grpo.GetByKey(docEntry))
                    {
                        company.GetLastError(
                            out int errorCode,
                            out string errorMessage);

                        Console.WriteLine("GRPO Load Failed");
                        Console.WriteLine(errorCode);
                        Console.WriteLine(errorMessage);

                        row.Status = "Load Failed";
                        continue;
                    }


                    Console.WriteLine("GRPO Loaded");
                    Console.WriteLine("Document Number : " + grpo.DocNum);


                    // Check if the GRPO is already Closed/Cancelled before attempting cancellation
                    if (grpo.DocumentStatus == BoStatus.bost_Close)
                    {
                        Console.WriteLine("GRPO is already Closed/Cancelled.");
                        row.Status = "Already Closed";
                        continue;
                    }


                    // Cancel GRPO
                    Documents cancelGRPO =
           grpo.CreateCancellationDocument();


                    int cancelResult = cancelGRPO.Add();


                    if (cancelResult == 0)
                    {
                        Console.WriteLine("GRPO Cancelled Successfully.");

                        int newDocEntry = Convert.ToInt32(company.GetNewObjectKey());

                        Console.WriteLine("Cancellation DocEntry : " + newDocEntry);

                        row.Status = "Closed";
                    }
                    else
                    {
                        company.GetLastError(
                            out int errorCode,
                            out string errorMessage);

                        Console.WriteLine("Cancellation Failed");
                        Console.WriteLine("Error Code : " + errorCode);
                        Console.WriteLine("Error Message : " + errorMessage);

                        row.Status = "Open";
                    }
                }

                // Write DocNum/Status log back into the same Excel file (Column B)
                WriteStatusToExcel(excelPath, grpoRows);

                // Export a DocNum,Status log file into the LOG folder
                ExportLogFile(grpoRows);
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


        //=================================================
        // Read all values in Column A (Sheet1) using OpenXML
        // until an empty/blank cell is found
        //=================================================
        static List<GrpoRow> ReadExcelColumnA(string filePath)
        {
            List<GrpoRow> docNums = new List<GrpoRow>();

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
                            docNums.Add(new GrpoRow
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


        //=================================================
        // Write Status (Closed / Open / Already Closed / Not Found / etc.)
        // back into Column B of the same Excel file, next to each DocNum
        //=================================================
        static void WriteStatusToExcel(string filePath, List<GrpoRow> grpoRows)
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

                    foreach (GrpoRow grpoRow in grpoRows)
                    {
                        uint rowIndex = (uint)grpoRow.RowIndex;

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
                            new CellValue(grpoRow.Status ?? string.Empty);
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


        //=================================================
        // Export DocNum,Status log to
        // D:\Navinesh\VS_Application\CloseGRPO_Reva\LOG
        //=================================================
        static void ExportLogFile(List<GrpoRow> grpoRows)
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
                    "GRPO_CancelLog_" +
                    DateTime.Now.ToString("yyyyMMdd_HHmmss") +
                    ".csv";

                string logFilePath = Path.Combine(logFolder, fileName);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("DocNum,Status");

                foreach (GrpoRow row in grpoRows)
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


        //=================================================
        // Extract column letters from a cell reference (e.g. "A23" -> "A")
        //=================================================
        static string GetColumnName(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return string.Empty;

            return new string(
                cellReference.TakeWhile(c => !char.IsDigit(c)).ToArray());
        }


        //=================================================
        // Get Excel Cell Value
        //=================================================
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
