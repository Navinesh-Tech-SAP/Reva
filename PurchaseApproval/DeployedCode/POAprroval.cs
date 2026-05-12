using System;
using Microsoft.Data.SqlClient;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Configuration;
using System.Collections.Generic;
using System.IO;

class Program
{
    static string logPath = ConfigurationManager.AppSettings["LogPath"];

    static void Main()
    {
        Log("==== APPLICATION START ====");

        try
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;
            int mailCount = 0;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                Log("Database connected successfully");

                string query = @"
            select 
                T1.email as ToEmail,
                T0.CreatorName,
                T2.email as CcEmail,
                T0.WebID as DraftNo,
                cast(T0.SendDate as date) as CreateDate
            from PURIND_H T0
            inner join [LIVE_REVA_University].dbo.OHEM T1 on T0.ApprovedBy = T1.empID
            inner join [LIVE_REVA_University].dbo.OHEM T2 on T0.CreatedBy = T2.empID
            where T0.ApprovalStatus = 'Pending' 
              and isnull(T0.email_A,'') = ''";

                SqlCommand cmd = new SqlCommand(query, con);
                SqlDataReader dr = cmd.ExecuteReader();

                var records = new List<(string DraftNo, string CreatorName, string CreateDate, string ToEmail, string CcEmail)>();

                while (dr.Read())
                {
                    records.Add((
                        dr["DraftNo"].ToString(),
                        dr["CreatorName"].ToString(),
                        Convert.ToDateTime(dr["CreateDate"]).ToString("dd-MM-yyyy"),
                        dr["ToEmail"].ToString(),
                        dr["CcEmail"].ToString()
                    ));
                }

                dr.Close();

                foreach (var rec in records)
                {
                    Log($"Processing DraftNo: {rec.DraftNo}");

                    if (string.IsNullOrWhiteSpace(rec.ToEmail))
                    {
                        Log($"Skipped DraftNo {rec.DraftNo} - No ToEmail");
                        continue;
                    }

                    string subject = $"Purchase Indent Draft No:{rec.DraftNo} Awaiting for Approval";

                    StringBuilder body = new StringBuilder();
                    body.Append("<p><strong>Dear Sir/Madam,</strong></p>");
                    body.Append("<p><strong>A new Draft Document has been awaiting for your approval. Kindly Approve the same.</strong></p>");

                    body.Append("<table border='1' cellpadding='5' cellspacing='0'>");
                    body.Append("<tr style='font-weight:bold;'>");
                    body.Append("<th>DRAFT NO.</th>");
                    body.Append("<th>APPROVAL STATUS</th>");
                    body.Append("<th>INDENT DATE</th>");
                    body.Append("<th>REQUESTER NAME</th>");
                    body.Append("</tr>");

                    body.Append("<tr>");
                    body.Append($"<td>{rec.DraftNo}</td>");
                    body.Append("<td>Document is Waiting for Approval</td>");
                    body.Append($"<td>{rec.CreateDate}</td>");
                    body.Append($"<td>{rec.CreatorName}</td>");
                    body.Append("</tr>");
                    body.Append("</table>");

                    bool isSent = SendMail(rec.ToEmail, rec.CcEmail, subject, body.ToString());

                    if (isSent)
                    {
                        using (SqlCommand updateCmd = new SqlCommand(
                            "UPDATE PURIND_H SET Email_A = 'Yes' WHERE WebID = @WebID", con))
                        {
                            updateCmd.Parameters.AddWithValue("@WebID", rec.DraftNo);
                            updateCmd.ExecuteNonQuery();
                        }

                        Log($"SUCCESS: Mail sent for DraftNo {rec.DraftNo}");
                        mailCount++;
                    }
                    else
                    {
                        Log($"FAILED: Mail not sent for DraftNo {rec.DraftNo}");
                    }
                }
            }

            Log($"Total Mails Sent: {mailCount}");
            Log("==== APPLICATION END SUCCESS ====");

            Environment.Exit(0); // ✅ success
        }
        catch (Exception ex)
        {
            Log("APPLICATION ERROR: " + ex.ToString());
            Log("==== APPLICATION END FAILURE ====");

            Environment.Exit(1); // ❌ failure
        }
    }

    static bool SendMail(string toEmail, string ccEmail, string subject, string body)
    {
        try
        {
            string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
            int port = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
            string fromEmail = ConfigurationManager.AppSettings["FromEmail"];
            string username = ConfigurationManager.AppSettings["Username"];
            string password = ConfigurationManager.AppSettings["Password"];

            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(fromEmail);
                mail.To.Add(toEmail);

                if (!string.IsNullOrWhiteSpace(ccEmail))
                    mail.CC.Add(ccEmail);

                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = true;

                using (SmtpClient smtp = new SmtpClient(smtpServer, port))
                {
                    smtp.EnableSsl = true;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(username, password);

                    smtp.Send(mail);
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            Log("MAIL ERROR: " + ex.Message + " | To: " + toEmail);
            return false;
        }
    }

    static void Log(string message)
    {
        try
        {
            if (!Directory.Exists(logPath))
                Directory.CreateDirectory(logPath);

            string filePath = Path.Combine(logPath, $"Log_{DateTime.Now:yyyyMMdd}.txt");

            using (StreamWriter sw = new StreamWriter(filePath, true))
            {
                sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
            }
        }
        catch
        {
            // avoid crash due to logging failure
        }
    }
}
