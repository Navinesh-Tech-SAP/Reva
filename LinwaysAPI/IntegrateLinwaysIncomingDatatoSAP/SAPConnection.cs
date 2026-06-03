using System;

class Program
{
    static void Main()
    {
        SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

        oCompany.Server = "SAPSERVER-1";
        oCompany.CompanyDB = "z_TEST_REVA_University";
        oCompany.UserName = "manager";
        oCompany.Password = "1234";

        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;

        oCompany.DbUserName = "sa";
        oCompany.DbPassword = "Welcome1#";

        oCompany.UseTrusted = false;
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;

        int result = oCompany.Connect();

        if (result == 0)
        {
            Console.WriteLine("SAP Connected Successfully!");
            Console.WriteLine("Company DB : " + oCompany.CompanyDB);
            Console.WriteLine("SAP User   : " + oCompany.UserName);

            oCompany.Disconnect();
        }
        else
        {
            int errCode;
            string errMsg;

            oCompany.GetLastError(out errCode, out errMsg);

            Console.WriteLine("SAP Connection Failed!");
            Console.WriteLine("Error Code : " + errCode);
            Console.WriteLine("Error Msg  : " + errMsg);
        }

        Console.ReadLine();
    }
}
