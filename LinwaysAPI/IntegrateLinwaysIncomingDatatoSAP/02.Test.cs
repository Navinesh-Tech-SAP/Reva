using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;

class Program
{
    static string sqlConnectionString =
        "Server=SAPSERVER-1;" +
        "Database=RevaIncomingPayment;" +
        "User Id=sa;" +
        "Password=Welcome1#;" +
        "TrustServerCertificate=True;";

    static string sapServer = "SAPSERVER-1";
    static string sapCompanyDB = "z_TEST_REVA_University";
    static string sapUserName = "manager";
    static string sapPassword = "1234";
    static string dbUserName = "sa";
    static string dbPassword = "Welcome1#";

    static string paymentAccount = "12090101";

    static void Main()
    {
        SAPbobsCOM.Company oCompany = null;

        try
        {
            oCompany = ConnectSAP();

            if (oCompany == null || !oCompany.Connected)
            {
                Console.WriteLine("SAP not connected.");
                return;
            }

            Console.WriteLine("SAP Connected Successfully.");

            EnsureUDF(oCompany, "ORCT", "REF_ID", "Linways Receipt No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            EnsureUDF(oCompany, "ORCT", "PymtType", "Payment Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            EnsureUDF(oCompany, "ORCT", "FeeName", "Fee Head Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            EnsureUDF(oCompany, "ORCT", "FeeHead", "Fee Head Amount", SAPbobsCOM.BoFieldTypes.db_Float, 0);

            using (SqlConnection conn = new SqlConnection(sqlConnectionString))
            {
                conn.Open();
                Console.WriteLine("SQL Connected Successfully.");

                List<ApiReceipt> receipts = GetPendingReceipts(conn);

                Console.WriteLine("Pending Records : " + receipts.Count);

                foreach (ApiReceipt receipt in receipts)
                {
                    try
                    {
                        if (receipt.IntegrationStatus == "Created")
                        {
                            Console.WriteLine("Already Created : " + receipt.ReceiptNo);
                            continue;
                        }

                        if (!BusinessPartnerExists(oCompany, receipt.AdmissionNo))
                        {
                            UpdateFailed(
                                conn,
                                receipt.ID,
                                "Customer Missing",
                                "Business Partner not found in SAP. CardCode: " + receipt.AdmissionNo
                            );

                            Console.WriteLine("BP Missing : " + receipt.AdmissionNo);
                            continue;
                        }

                        if (IncomingPaymentAlreadyExists(oCompany, receipt.ReceiptNo))
                        {
                            UpdateFailed(
                                conn,
                                receipt.ID,
                                "Created",
                                "Incoming Payment already exists in SAP for ReceiptNo: " + receipt.ReceiptNo
                            );

                            Console.WriteLine("Already Exists in SAP : " + receipt.ReceiptNo);
                            continue;
                        }

                        CreateIncomingPayment(oCompany, conn, receipt);

                        Console.WriteLine("Created Incoming Payment : " + receipt.ReceiptNo);
                    }
                    catch (Exception ex)
                    {
                        UpdateFailed(conn, receipt.ID, "Failed", ex.Message);
                        Console.WriteLine("Failed : " + receipt.ReceiptNo + " - " + ex.Message);
                    }
                }
            }

            Console.WriteLine("SAP Integration Completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error:");
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (oCompany != null && oCompany.Connected)
                oCompany.Disconnect();

            if (oCompany != null)
                Marshal.ReleaseComObject(oCompany);
        }

        Console.ReadLine();
    }

    static SAPbobsCOM.Company ConnectSAP()
    {
        SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

        oCompany.Server = sapServer;
        oCompany.CompanyDB = sapCompanyDB;
        oCompany.UserName = sapUserName;
        oCompany.Password = sapPassword;

        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
        oCompany.DbUserName = dbUserName;
        oCompany.DbPassword = dbPassword;

        oCompany.UseTrusted = false;
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;

        int result = oCompany.Connect();

        if (result != 0)
        {
            oCompany.GetLastError(out int errCode, out string errMsg);
            Console.WriteLine("SAP Connection Failed.");
            Console.WriteLine("Error Code : " + errCode);
            Console.WriteLine("Error Msg  : " + errMsg);
            return null;
        }

        return oCompany;
    }

    static List<ApiReceipt> GetPendingReceipts(SqlConnection conn)
    {
        List<ApiReceipt> list = new List<ApiReceipt>();

        string sql = @"
SELECT 
    ID,
    AdmissionNo,
    StudentName,
    ReceiptNo,
    ReceiptDate,
    PaymentType,
    FeeHeadName,
    FeeHeadAmount,
    TotalAmount,
    ISNULL(IntegrationStatus, 'Pending') AS IntegrationStatus
FROM API_StudentFeeReceipt
WHERE ISNULL(IntegrationStatus, 'Pending') <> 'Created'
ORDER BY ReceiptDate, ID";

        using (SqlCommand cmd = new SqlCommand(sql, conn))
        using (SqlDataReader dr = cmd.ExecuteReader())
        {
            while (dr.Read())
            {
                list.Add(new ApiReceipt
                {
                    ID = Convert.ToInt32(dr["ID"]),
                    AdmissionNo = Convert.ToString(dr["AdmissionNo"]),
                    StudentName = Convert.ToString(dr["StudentName"]),
                    ReceiptNo = Convert.ToString(dr["ReceiptNo"]),
                    ReceiptDate = Convert.ToDateTime(dr["ReceiptDate"]),
                    PaymentType = Convert.ToString(dr["PaymentType"]),
                    FeeHeadName = Convert.ToString(dr["FeeHeadName"]),
                    FeeHeadAmount = Convert.ToDecimal(dr["FeeHeadAmount"]),
                    TotalAmount = Convert.ToDecimal(dr["TotalAmount"]),
                    IntegrationStatus = Convert.ToString(dr["IntegrationStatus"])
                });
            }
        }

        return list;
    }

    static bool BusinessPartnerExists(SAPbobsCOM.Company oCompany, string cardCode)
    {
        SAPbobsCOM.BusinessPartners bp = null;

        try
        {
            bp = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.oBusinessPartners);

            return bp.GetByKey(cardCode);
        }
        finally
        {
            if (bp != null)
                Marshal.ReleaseComObject(bp);
        }
    }

    static bool IncomingPaymentAlreadyExists(SAPbobsCOM.Company oCompany, string receiptNo)
    {
        SAPbobsCOM.Recordset rs = null;

        try
        {
            rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset);

            string safeReceiptNo = receiptNo.Replace("'", "''");

            string sql = $@"
SELECT DocEntry 
FROM ORCT 
WHERE U_REF_ID = '{safeReceiptNo}'
AND Canceled = 'N'";

            rs.DoQuery(sql);

            return !rs.EoF;
        }
        finally
        {
            if (rs != null)
                Marshal.ReleaseComObject(rs);
        }
    }

    static void CreateIncomingPayment(
        SAPbobsCOM.Company oCompany,
        SqlConnection conn,
        ApiReceipt receipt)
    {
        SAPbobsCOM.Payments payment = null;

        try
        {
            payment = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.oIncomingPayments);

            payment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer;
            payment.CardCode = receipt.AdmissionNo;
            payment.BPLID = 3;

            //payment.DocCurrency = "INR";

            payment.DocDate = receipt.ReceiptDate;
            payment.TaxDate = receipt.ReceiptDate;
            payment.DueDate = receipt.ReceiptDate;

            payment.Remarks = "Linways Receipt No: " + receipt.ReceiptNo;
            payment.JournalRemarks = "Linways Incoming Payment";

            string payType = receipt.PaymentType.ToUpper();

            if (payType == "CASH")
            {
                payment.CashAccount = paymentAccount;
                payment.CashSum = Convert.ToDouble(receipt.TotalAmount);
            }
            else
            {
                payment.TransferAccount = "12100209";
                payment.TransferSum = Convert.ToDouble(receipt.TotalAmount);
                payment.TransferDate = receipt.ReceiptDate;
                payment.TransferReference = receipt.ReceiptNo;
            }

            payment.UserFields.Fields.Item("U_REF_ID").Value = receipt.ReceiptNo;
            payment.UserFields.Fields.Item("U_PymtType").Value = receipt.PaymentType;
            payment.UserFields.Fields.Item("U_FeeName").Value = receipt.FeeHeadName;
            payment.UserFields.Fields.Item("U_FeeHead").Value = Convert.ToDouble(receipt.FeeHeadAmount);
            Console.WriteLine("CardCode      : " + payment.CardCode);
            Console.WriteLine("Branch        : " + payment.BPLID);
            Console.WriteLine("DocCurrency   : " + payment.DocCurrency);
            Console.WriteLine("Posting Date  : " + payment.DocDate);
            Console.WriteLine("CashSum       : " + payment.CashSum);
            Console.WriteLine("TransferSum   : " + payment.TransferSum);

            int result = payment.Add();


            if (result != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                throw new Exception(errCode + " - " + errMsg);
            }

            string newDocEntry = oCompany.GetNewObjectKey();
            int docEntry = Convert.ToInt32(newDocEntry);
            int docNum = GetIncomingPaymentDocNum(oCompany, docEntry);

            UpdateSuccess(conn, receipt.ID, docEntry, docNum);
        }
        finally
        {
            if (payment != null)
                Marshal.ReleaseComObject(payment);
        }
    }

    static int GetIncomingPaymentDocNum(SAPbobsCOM.Company oCompany, int docEntry)
    {
        SAPbobsCOM.Recordset rs = null;

        try
        {
            rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery("SELECT DocNum FROM ORCT WHERE DocEntry = " + docEntry);

            if (!rs.EoF)
                return Convert.ToInt32(rs.Fields.Item("DocNum").Value);

            return 0;
        }
        finally
        {
            if (rs != null)
                Marshal.ReleaseComObject(rs);
        }
    }

    static void UpdateSuccess(SqlConnection conn, int id, int sapDocEntry, int sapDocNum)
    {
        string sql = @"
UPDATE API_StudentFeeReceipt
SET
    SAPDocEntry = @SAPDocEntry,
    SAPDocNum = @SAPDocNum,
    IntegrationStatus = 'Created',
    IntegrationMessage = 'Incoming Payment Created Successfully',
    UpdatedDate = GETDATE()
WHERE ID = @ID";

        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            cmd.Parameters.AddWithValue("@SAPDocEntry", sapDocEntry);
            cmd.Parameters.AddWithValue("@SAPDocNum", sapDocNum);
            cmd.Parameters.AddWithValue("@ID", id);
            cmd.ExecuteNonQuery();
        }
    }

    static void UpdateFailed(SqlConnection conn, int id, string status, string message)
    {
        string sql = @"
UPDATE API_StudentFeeReceipt
SET
    IntegrationStatus = @IntegrationStatus,
    IntegrationMessage = @IntegrationMessage,
    UpdatedDate = GETDATE()
WHERE ID = @ID";

        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            cmd.Parameters.AddWithValue("@IntegrationStatus", status);
            cmd.Parameters.AddWithValue("@IntegrationMessage", message ?? "");
            cmd.Parameters.AddWithValue("@ID", id);
            cmd.ExecuteNonQuery();
        }
    }

    static void EnsureUDF(
        SAPbobsCOM.Company oCompany,
        string tableName,
        string fieldName,
        string description,
        SAPbobsCOM.BoFieldTypes fieldType,
        int size)
    {
        if (UDFExists(oCompany, tableName, fieldName))
        {
            Console.WriteLine("UDF already exists : U_" + fieldName);
            return;
        }

        SAPbobsCOM.UserFieldsMD udf = null;

        try
        {
            udf = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.oUserFields);

            udf.TableName = tableName;
            udf.Name = fieldName;
            udf.Description = description;
            udf.Type = fieldType;

            if (fieldType == SAPbobsCOM.BoFieldTypes.db_Alpha)
            {
                udf.Size = size;
            }

            if (fieldType == SAPbobsCOM.BoFieldTypes.db_Float)
            {
                udf.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum;
            }

            int result = udf.Add();

            if (result != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                throw new Exception("UDF Create Failed U_" + fieldName + " : " + errCode + " - " + errMsg);
            }

            Console.WriteLine("UDF created : U_" + fieldName);
        }
        finally
        {
            if (udf != null)
                Marshal.ReleaseComObject(udf);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    static bool UDFExists(SAPbobsCOM.Company oCompany, string tableName, string fieldName)
    {
        SAPbobsCOM.Recordset rs = null;

        try
        {
            rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sql = $@"
SELECT FieldID
FROM CUFD
WHERE TableID = '{tableName}'
AND AliasID = '{fieldName}'";

            rs.DoQuery(sql);

            return !rs.EoF;
        }
        finally
        {
            if (rs != null)
                Marshal.ReleaseComObject(rs);
        }
    }
}

public class ApiReceipt
{
    public int ID { get; set; }
    public string AdmissionNo { get; set; }
    public string StudentName { get; set; }
    public string ReceiptNo { get; set; }
    public DateTime ReceiptDate { get; set; }
    public string PaymentType { get; set; }
    public string FeeHeadName { get; set; }
    public decimal FeeHeadAmount { get; set; }
    public decimal TotalAmount { get; set; }
    public string IntegrationStatus { get; set; }
}
