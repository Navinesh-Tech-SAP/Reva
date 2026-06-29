using System;
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
        company.UserName = "manager";      // SAP User
        company.Password = "1234";      // SAP Password

        company.DbServerType = BoDataServerTypes.dst_MSSQL2019;
        company.language = BoSuppLangs.ln_English;


        int result = company.Connect();

        if (result == 0)
        {
            Console.WriteLine("SAP Business One Connected Successfully!");
            Console.WriteLine("Company Name: " + company.CompanyName);
            Console.WriteLine("Database: " + company.CompanyDB);
        }
        else
        {
            company.GetLastError(out int errorCode, out string errorMessage);

            Console.WriteLine("Connection Failed!");
            Console.WriteLine("Error Code: " + errorCode);
            Console.WriteLine("Error Message: " + errorMessage);
        }

        Console.ReadLine();
    }
}
