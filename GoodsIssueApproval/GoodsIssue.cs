
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

class Program
{
    static string logPath = ConfigurationManager.AppSettings["LogPath"];

    static void Main()
    {
        Log("==================================================");
        Log("APPLICATION START");

        try
        {
            string connectionString =
                ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;

            int mailCount = 0;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                Log("Database connected successfully");

                string query = @"
                    SELECT 
                        T1.email AS ToEmail,
                        T0.CreatorName,
                        T2.email AS CcEmail,
                        T0.WebID AS DraftNo,
                        CAST(T0.SendDate AS DATE) AS CreateDate,
                        T0.Email_A AS EmailSent
                    FROM GOODS_ISSUE_H T0
                    INNER JOIN [LIVE_REVA_University].dbo.OHEM T1 
                        ON T0.ApprovedBy = T1.empID
                    INNER JOIN [LIVE_REVA_University].dbo.OHEM T2 
                        ON T0.CreatedBy = T2.empID
                    WHERE T0.ApprovalStatus = 'Pending'
                    AND ISNULL(T0.Email_A,'') = ''";

                SqlCommand cmd = new SqlCommand(query, con);

                SqlDataReader dr = cmd.ExecuteReader();

                var records = new List<(
                    string DraftNo,
                    string CreatorName,
                    string CreateDate,
                    string ToEmail,
                    string CcEmail
                )>();

                while (dr.Read())
                {
                    records.Add((
                        dr["DraftNo"]?.ToString() ?? "",
                        dr["CreatorName"]?.ToString() ?? "",
                        Convert.ToDateTime(dr["CreateDate"]).ToString("dd-MM-yyyy"),
                        dr["ToEmail"]?.ToString() ?? "",
                        dr["CcEmail"]?.ToString() ?? ""
                    ));
                }

                dr.Close();

                Log($"Total Pending Records Found : {records.Count}");

                foreach (var rec in records)
                {
                    try
                    {
                        Log($"--------------------------------------------------");
                        Log($"Processing Draft No : {rec.DraftNo}");

                        if (string.IsNullOrWhiteSpace(rec.ToEmail))
                        {
                            Log($"FAILED : ToEmail is empty for Draft No : {rec.DraftNo}");
                            continue;
                        }

                        string subject =
                            $"Goods Issue Draft No:{rec.DraftNo} Awaiting for Approval";

                        StringBuilder body = new StringBuilder();

                        body.Append("<p><strong>Dear Sir/Madam,</strong></p>");

                        body.Append("<p><strong>");
                        body.Append("A new Draft Document has been awaiting for your approval. Kindly Approve the same.");
                        body.Append("</strong></p>");

                        body.Append("<table border='1' cellpadding='5' cellspacing='0'>");

                        body.Append("<tr style='font-weight:bold;'>");
                        body.Append("<th>DRAFT NO.</th>");
                        body.Append("<th>APPROVAL STATUS</th>");
                        body.Append("<th>GOODS ISSUE DATE</th>");
                        body.Append("<th>REQUESTER NAME</th>");
                        body.Append("</tr>");

                        body.Append("<tr>");
                        body.Append($"<td>{rec.DraftNo}</td>");
                        body.Append("<td>Document is Waiting for Approval</td>");
                        body.Append($"<td>{rec.CreateDate}</td>");
                        body.Append($"<td>{rec.CreatorName}</td>");
                        body.Append("</tr>");

                        body.Append("</table>");

                        bool isSent = SendMail(
                            rec.ToEmail,
                            rec.CcEmail,
                            subject,
                            body.ToString(),
                            rec.DraftNo
                        );

                        if (isSent)
                        {
                            try
                            {
                                using (SqlCommand updateCmd = new SqlCommand(
                                    "UPDATE GOODS_ISSUE_H SET Email_A = 'Yes' WHERE WebID = @WebID",
                                    con))
                                {
                                    updateCmd.Parameters.AddWithValue("@WebID", rec.DraftNo);

                                    int rows = updateCmd.ExecuteNonQuery();

                                    if (rows > 0)
                                    {
                                        Log($"Database updated successfully for Draft No : {rec.DraftNo}");
                                    }
                                    else
                                    {
                                        Log($"WARNING : No rows updated for Draft No : {rec.DraftNo}");
                                    }
                                }

                                Log($"SUCCESS : Mail sent successfully for Draft No : {rec.DraftNo}");

                                mailCount++;
                            }
                            catch (Exception dbEx)
                            {
                                Log($"DATABASE UPDATE ERROR for Draft No : {rec.DraftNo}");
                                Log($"ERROR MESSAGE : {dbEx.Message}");
                                Log($"STACK TRACE : {dbEx.StackTrace}");
                            }
                        }
                        else
                        {
                            Log($"FAILED : Mail sending failed for Draft No : {rec.DraftNo}");
                        }
                    }
                    catch (Exception loopEx)
                    {
                        Log($"PROCESSING ERROR for Draft No : {rec.DraftNo}");
                        Log($"ERROR MESSAGE : {loopEx.Message}");
                        Log($"STACK TRACE : {loopEx.StackTrace}");
                    }
                }
            }

            Log($"Total Mails Sent Successfully : {mailCount}");

            Log("APPLICATION END SUCCESS");
            Log("==================================================");

            Environment.Exit(0);
        }
        catch (Exception ex)
        {
            Log("FATAL APPLICATION ERROR");
            Log($"ERROR MESSAGE : {ex.Message}");
            Log($"STACK TRACE : {ex.StackTrace}");
            Log("APPLICATION END FAILURE");
            Log("==================================================");

            Environment.Exit(1);
        }
    }

    static bool SendMail(
        string toEmail,
        string ccEmail,
        string subject,
        string body,
        string draftNo)
    {
        try
        {
            string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];

            int port =
                int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);

            string fromEmail =
                ConfigurationManager.AppSettings["FromEmail"];

            string username =
                ConfigurationManager.AppSettings["Username"];

            string password =
                ConfigurationManager.AppSettings["Password"];

            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(fromEmail);

                mail.To.Add(toEmail);

                if (!string.IsNullOrWhiteSpace(ccEmail))
                {
                    mail.CC.Add(ccEmail);
                }

                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = true;

                using (SmtpClient smtp = new SmtpClient(smtpServer, port))
                {
                    smtp.EnableSsl = true;
                    smtp.UseDefaultCredentials = false;

                    smtp.Credentials =
                        new NetworkCredential(username, password);

                    smtp.Send(mail);
                }
            }

            Log($"MAIL SUCCESS : Draft No : {draftNo} | To : {toEmail}");

            return true;
        }
        catch (SmtpException smtpEx)
        {
            Log($"SMTP ERROR for Draft No : {draftNo}");
            Log($"SMTP STATUS CODE : {smtpEx.StatusCode}");
            Log($"ERROR MESSAGE : {smtpEx.Message}");
            Log($"STACK TRACE : {smtpEx.StackTrace}");

            return false;
        }
        catch (Exception ex)
        {
            Log($"MAIL ERROR for Draft No : {draftNo}");
            Log($"ERROR MESSAGE : {ex.Message}");
            Log($"STACK TRACE : {ex.StackTrace}");

            return false;
        }
    }

    static void Log(string message)
    {
        try
        {
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            string filePath =
                Path.Combine(
                    logPath,
                    $"Log_{DateTime.Now:yyyyMMdd}.txt");

            using (StreamWriter sw = new StreamWriter(filePath, true))
            {
                sw.WriteLine(
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
            }
        }
        catch
        {
            // Avoid application crash due to logging issue
        }
    }
}
