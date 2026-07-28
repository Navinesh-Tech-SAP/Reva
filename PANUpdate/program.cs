using System;
using SAPbobsCOM;

namespace UpdatePAN
{
    class Program
    {
        static void Main(string[] args)
        {
            Company oCompany = new Company();

            oCompany.Server = "SAPSERVER-1";
            oCompany.CompanyDB = "REVA_LIVE";
            oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019;
            oCompany.DbUserName = "sa";
            oCompany.DbPassword = "Welcome1#";
            oCompany.UserName = "manager";
            oCompany.Password = "1234";
            oCompany.LicenseServer = "SAPSERVER-1:30000";

            int ret = oCompany.Connect();

            if (ret != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                Console.WriteLine($"Connection Failed : {errCode} - {errMsg}");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Connected Successfully.");

            BusinessPartners bp =
                (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            string cardCode = "V07153L";
            //string cardCode = "V06986L";

            if (!bp.GetByKey(cardCode))
            {
                Console.WriteLine("Business Partner not found.");
                Console.ReadKey();
                return;
            }

            // Read PAN from UDF
            string pan = Convert.ToString(bp.UserFields.Fields.Item("U_PANCARDNO").Value);

            Console.WriteLine("PAN From UDF : " + pan);

            if (string.IsNullOrWhiteSpace(pan))
            {
                Console.WriteLine("U_PANCARDNO is empty.");
                Console.ReadKey();
                return;
            }

            // Update PAN Number
            bp.FiscalTaxID.TaxId0 = pan;

            ret = bp.Update();

            if (ret == 0)
            {
                Console.WriteLine("PAN Number Updated Successfully.");
            }
            else
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                Console.WriteLine($"Update Failed : {errCode} - {errMsg}");
            }

            oCompany.Disconnect();

            Console.ReadKey();
        }
    }
}
