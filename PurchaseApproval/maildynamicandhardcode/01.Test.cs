using System;
using Microsoft.Data.SqlClient;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Configuration;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        try
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;

            int mailCount = 0;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

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
                where ApprovalStatus = 'Pending' 
                  and isnull(T0.email_A,'') = ''";

                SqlCommand cmd = new SqlCommand(query, con);
                SqlDataReader dr = cmd.ExecuteReader();

                // =====  LIVE (Dynamic Emails from DB) =====
                var records = new List<(string DraftNo, string CreatorName, string CreateDate, string ToEmail, string CcEmail)>();

                while (dr.Read())
                {
                    records.Add((
                        dr["DraftNo"].ToString(),
                        dr["CreatorName"].ToString(),
                        Convert.ToDateTime(dr["CreateDate"]).ToString("dd-MM-yyyy"),
                        dr["ToEmail"].ToString(),   // <-- LIVE
                        dr["CcEmail"].ToString()    // <-- LIVE
                    ));
                }

                dr.Close();

                // =====  TEST (Hardcoded Emails) =====
                //  Uncomment below lines for TEST
                string testToEmail = "navineshargit@gmail.com";
                string testCcEmail = "navineshn103@gmail.com";

                foreach (var rec in records)
                {
                    // ===== SWITCH HERE =====

                    //  TEST MODE (currently active)
                    string toEmail = testToEmail;
                    string ccEmail = testCcEmail;

                    //  LIVE MODE (use below instead)
                    // string toEmail = rec.ToEmail;
                    // string ccEmail = rec.CcEmail;

                    if (string.IsNullOrWhiteSpace(toEmail))
                    {
                        Console.WriteLine($"Skipped Draft No {rec.DraftNo} (No ToEmail)");
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

                    body.Append("<br/><p>Regards,</p>");
                    body.Append("<p>REVA Purchase Team<br/>");
                    body.Append("REVA UNIVERSITY | Rukmini Knowledge Park | Kattigenahalli | Yelahanka | Bengaluru | Karnataka 560 064</p>");

                    bool isSent = SendMail(toEmail, ccEmail, subject, body.ToString());

                    //if (isSent)
                    //{
                    //    using (SqlCommand updateCmd = new SqlCommand(
                    //        "UPDATE PURIND_H SET Email_A = 'Y' WHERE WebID = @WebID", con))
                    //    {
                    //        updateCmd.Parameters.AddWithValue("@WebID", rec.DraftNo);
                    //        updateCmd.ExecuteNonQuery();
                    //    }

                    //    Console.WriteLine($"Mail sent & updated for Draft No: {rec.DraftNo}");
                    //    mailCount++;
                    //}
                    //else
                    //{
                    //    Console.WriteLine($"Mail failed for Draft No: {rec.DraftNo}");
                    //}
                }
            }

            Console.WriteLine($"Total Mails Sent: {mailCount}");
            Environment.Exit(0);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Application Error: " + ex.Message);
            Environment.Exit(1);
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
            Console.WriteLine("Mail Error: " + ex.Message);
            return false;
        }
    }
}
