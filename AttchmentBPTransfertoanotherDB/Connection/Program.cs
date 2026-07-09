using System;
using SAPbobsCOM;

namespace SAPConnection
{
    class Program
    {
        static void Main(string[] args)
        {
            Company company = new Company();

            company.Server = "SAPSERVER-1";
            company.CompanyDB = "REVA_LIVE";
            company.DbServerType = BoDataServerTypes.dst_MSSQL2019; // Change if needed
            company.DbUserName = "sa";
            company.DbPassword = "Welcome1#";

            company.UserName = "manager";
            company.Password = "1234";

            company.LicenseServer = "SAPSERVER-1:30000";

            int ret = company.Connect();

            if (ret != 0)
            {
                int errCode;
                string errMsg;

                company.GetLastError(out errCode, out errMsg);
                Console.WriteLine("Connection failed: " + errCode + " - " + errMsg);
            }
            else
            {
                Console.WriteLine("Connected successfully.");
            }

            Console.ReadLine();
        }
    }
}
