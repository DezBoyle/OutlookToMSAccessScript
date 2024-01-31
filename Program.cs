﻿using System.Runtime.Versioning;

namespace OutlookToMSAccessScript
{
    [SupportedOSPlatform("windows")]
    internal class Program
    {
        static void Main(string[] args)
        {
            AccessDatabaseTool accessDatabaseTool = new AccessDatabaseTool("C:/Users/dboyle/testdb.mdb");

            while(true)
            {
                Console.WriteLine("(TEST) Enter Company Name:");
                string companyName = Console.ReadLine();

                bool companyExists = accessDatabaseTool.RowExists("CompanyName", "CompanyName", companyName);
                Console.WriteLine("Company Exists: " + companyExists);
                if(!companyExists)
                {
                    Console.WriteLine("new company- Added company to table");
                    accessDatabaseTool.AddRow("CompanyName", "CompanyName", companyName);
                }

                //Add/Update information to company row
                Console.WriteLine($"(TEST) Enter Company {companyName} Address:");
                string address = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} City:");
                string city = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} State:");
                string state = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} Zip:");
                string zip = Console.ReadLine();

                KeyValuePair<string, string>[] properties = new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("COAddress", address),
                    new KeyValuePair<string, string>("COCity", city),
                    new KeyValuePair<string, string>("COState", state),
                    new KeyValuePair<string, string>("COZip", zip),
                };

                accessDatabaseTool.UpdateRow("CompanyName", "CompanyName", companyName, properties);

                Console.WriteLine($"(TEST) Enter Company {companyName} First Name:");
                string firstName = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} Last Name:");
                string lastName = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} Email:");
                string email = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} Phone:");
                string phone = Console.ReadLine();

                string companyId = accessDatabaseTool.GetRows("CompanyName", "CompanyName", companyName).Rows[0].ToString(); //grab the company id (may have multiple, just take the first one)

                bool contactExists = accessDatabaseTool.RowExists("Contact information", "COCompany", companyId);
                //same thing as before, make new if need to, then update it with info
            }
           
        }
    }
}