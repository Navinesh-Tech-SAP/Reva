using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Configuration;

class Program
{
    static readonly string connectionString =
     ConfigurationManager.AppSettings["ConnectionString"];

    static readonly string baseUrl =
        ConfigurationManager.AppSettings["BaseUrl"];

    static readonly string apiUserName =
        ConfigurationManager.AppSettings["ApiUserName"];

    static readonly string apiPassword =
        ConfigurationManager.AppSettings["ApiPassword"];

    static readonly string apiKey =
        ConfigurationManager.AppSettings["ApiKey"];

    static readonly string apiSecretKey =
        ConfigurationManager.AppSettings["ApiSecretKey"];

    static readonly string defaultStartDate =
        ConfigurationManager.AppSettings["DefaultStartDate"];
    static string[] paymentMethods =
 {
    "CASH",
    "DD",
    "CHEQUE",
    "POS",
    "RTGS",
    "NEFT",
    "UPI",
    "ONLINE",
    "Wallet"
};

    static async Task Main()
    {
        try
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                Console.WriteLine("SQL Connected Successfully.");

              
                using (HttpClient client = new HttpClient())
                {
                    SetAuthentication(client);

                    DateTime startDate =
        DateTime.ParseExact(
            GetFromDate(conn),
            "dd-MM-yyyy",
            CultureInfo.InvariantCulture);

                    DateTime endDate = DateTime.Now.AddDays(-1);

                    Console.WriteLine("Start Date : " + startDate.ToString("dd-MM-yyyy"));
                    Console.WriteLine("End Date   : " + endDate.ToString("dd-MM-yyyy"));

                    while (startDate <= endDate)
                    {
                        DateTime chunkEndDate = startDate.AddDays(29);

                        if (chunkEndDate > endDate)
                            chunkEndDate = endDate;

                        string fromDate = startDate.ToString("dd-MM-yyyy");
                        string toDate = chunkEndDate.ToString("dd-MM-yyyy");

                        Console.WriteLine("--------------------------------");
                        Console.WriteLine($"Date Range : {fromDate} To {toDate}");

                        foreach (string paymentMethod in paymentMethods)
                        {
                            Console.WriteLine("--------------------------------");
                            Console.WriteLine("Processing Payment Method : " + paymentMethod);

                            string url =
                                baseUrl +
                                "?fromDate=" + fromDate +
                                "&toDate=" + toDate +
                                "&accountId=1" +
                                "&paymentMethod=" + Uri.EscapeDataString(paymentMethod);

                            HttpResponseMessage response =
                                await client.GetAsync(url);

                            if (!response.IsSuccessStatusCode)
                            {
                                Console.WriteLine("API Failed : " + paymentMethod);
                                Console.WriteLine("Status Code: " + response.StatusCode);
                                continue;
                            }

                            string json = await response.Content.ReadAsStringAsync();

                            ApiResponse apiResponse =
                                JsonConvert.DeserializeObject<ApiResponse>(json);

                            if (apiResponse == null ||
                                !apiResponse.success ||
                                apiResponse.data == null)
                            {
                                Console.WriteLine("No data found for " + paymentMethod);
                                continue;
                            }

                            int insertedCount = 0;
                            int duplicateCount = 0;

                            foreach (StudentFeeReceipt receipt in apiResponse.data)
                            {
                                if (receipt.feeHeads == null ||
                                    receipt.feeHeads.Count == 0)
                                    continue;

                                foreach (FeeHead feeHead in receipt.feeHeads)
                                {
                                    if (IsReceiptExists(
                                        conn,
                                        receipt.receiptNo,
                                        feeHead.name))
                                    {
                                        duplicateCount++;
                                        continue;
                                    }

                                    InsertReceipt(conn, receipt, feeHead);
                                    insertedCount++;
                                }
                            }

                            Console.WriteLine("Inserted  : " + insertedCount);
                            Console.WriteLine("Duplicate : " + duplicateCount);
                        }

                        startDate = chunkEndDate.AddDays(1);
                    }
                }

                Console.WriteLine("--------------------------------");
                Console.WriteLine("API Import Completed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error:");
            Console.WriteLine(ex.Message);
        }

        Console.ReadLine();
    }

    static void SetAuthentication(HttpClient client)
    {
        string authValue = Convert.ToBase64String(
            Encoding.ASCII.GetBytes(apiUserName + ":" + apiPassword)
        );

        client.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Basic", authValue);

        client.DefaultRequestHeaders.Add("apiKey", apiKey);
        client.DefaultRequestHeaders.Add("apiSecretKey", apiSecretKey);
    }

    static string GetFromDate(SqlConnection conn)
    {
        string sql = @"
SELECT ISNULL(
    CONVERT(VARCHAR(10), DATEADD(DAY, 1, MAX(ReceiptDate)), 105),
    '01-01-2026'
)
FROM API_StudentFeeReceipt";

        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            return Convert.ToString(cmd.ExecuteScalar());
        }
    }

    static bool IsReceiptExists(SqlConnection conn, string receiptNo, string feeHeadName)
    {
        string sql = @"
SELECT COUNT(*)
FROM API_StudentFeeReceipt
WHERE ReceiptNo = @ReceiptNo
AND FeeHeadName = @FeeHeadName";

        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            cmd.Parameters.AddWithValue("@ReceiptNo", receiptNo ?? "");
            cmd.Parameters.AddWithValue("@FeeHeadName", feeHeadName ?? "");

            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }
    }

    static void InsertReceipt(SqlConnection conn, StudentFeeReceipt receipt, FeeHead feeHead)
    {
        string sql = @"
INSERT INTO API_StudentFeeReceipt
(
    AdmissionNo,
    StudentName,
    ReceiptNo,
    ReceiptDate,
    PaymentType,
    FeeHeadName,
    FeeHeadAmount,
    TotalAmount,
    IntegrationStatus
)
VALUES
(
    @AdmissionNo,
    @StudentName,
    @ReceiptNo,
    @ReceiptDate,
    @PaymentType,
    @FeeHeadName,
    @FeeHeadAmount,
    @TotalAmount,
    'Pending'
)";

        DateTime receiptDate =
            DateTime.ParseExact(receipt.date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

        decimal feeAmount = 0;
        decimal.TryParse(feeHead.amount, out feeAmount);

        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            cmd.Parameters.AddWithValue("@AdmissionNo", receipt.admissionNo ?? "");
            cmd.Parameters.AddWithValue("@StudentName", receipt.studentName ?? "");
            cmd.Parameters.AddWithValue("@ReceiptNo", receipt.receiptNo ?? "");
            cmd.Parameters.AddWithValue("@ReceiptDate", receiptDate);
            cmd.Parameters.AddWithValue("@PaymentType", receipt.type ?? "");
            cmd.Parameters.AddWithValue("@FeeHeadName", feeHead.name ?? "");
            cmd.Parameters.AddWithValue("@FeeHeadAmount", feeAmount);
            cmd.Parameters.AddWithValue("@TotalAmount", receipt.totalAmount);

            cmd.ExecuteNonQuery();
        }
    }
}

public class ApiResponse
{
    public string id { get; set; }
    public bool success { get; set; }
    public List<StudentFeeReceipt> data { get; set; }
}

public class StudentFeeReceipt
{
    public string admissionNo { get; set; }
    public string studentName { get; set; }
    public string receiptNo { get; set; }
    public string date { get; set; }
    public string type { get; set; }
    public decimal totalAmount { get; set; }
    public List<FeeHead> feeHeads { get; set; }
}

public class FeeHead
{
    public string name { get; set; }
    public string amount { get; set; }
}
