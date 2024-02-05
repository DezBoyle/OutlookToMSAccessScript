using System.Runtime.Versioning;

namespace OutlookToMSAccessScript
{
    [SupportedOSPlatform("windows")]
    internal class Program
    {
        static void Main(string[] args)
        {
            string saveFilePath = "databasePath.txt";
            string databasePath = "C:/Users/dboyle/testdb.mdb";

            if(File.Exists(saveFilePath))
            { databasePath = File.ReadAllText(saveFilePath); }
            else
            {
                Console.WriteLine("Enter the path of the database file (.mdb)");
                File.WriteAllText(saveFilePath, Console.ReadLine());
            }

            File.WriteAllText("information.txt", "Created by Dez Boyle\nSource Code: https://github.com/DezBoyle/OutlookToMSAccessScript");

            DebugPrompt(databasePath);
        }

        private static void DebugPrompt(string databasePath)
        {
            AccessDatabaseTool accessDatabaseTool = new AccessDatabaseTool(databasePath);

            while (true)
            {
                Console.WriteLine("(TEST) Enter Company Name:");
                string companyName = Console.ReadLine();

                bool companyExists = accessDatabaseTool.RowExists("CompanyName", "CompanyName", companyName);
                Console.WriteLine("Company Exists: " + companyExists);
                if (!companyExists)
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

                KeyValuePair<string, string>[] companyProperties = new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("COAddress", address),
                    new KeyValuePair<string, string>("COCity", city),
                    new KeyValuePair<string, string>("COState", state),
                    new KeyValuePair<string, string>("COZip", zip),
                };

                accessDatabaseTool.UpdateRow("CompanyName", "CompanyName", $"'{companyName}'", companyProperties);

                Console.WriteLine($"(TEST) Enter Company {companyName} First Name:");
                string firstName = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} Last Name:");
                string lastName = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} Email:");
                string email = Console.ReadLine();
                Console.WriteLine($"(TEST) Enter Company {companyName} Phone (FORMAT: 8154425564):");
                string phone = Console.ReadLine();

                string companyId = accessDatabaseTool.GetRows("CompanyName", "CompanyName", companyName).Rows[0][0].ToString(); //grab the company id (may have multiple, just take the first one)

                bool contactExists = accessDatabaseTool.RowExists("Contact information", "COCompany", companyId);
                if (!contactExists)
                {
                    Console.WriteLine("new contact- Added contact to table");
                    accessDatabaseTool.AddRow("Contact information", "COCompanyID", companyId);
                }

                KeyValuePair<string, string>[] contactProperties = new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("ContactFName", firstName),
                    new KeyValuePair<string, string>("ContactLName", lastName),
                    new KeyValuePair<string, string>("ContactEmail", email),
                    new KeyValuePair<string, string>("ContractPhone", phone), //they spelled Contact wrong in the database lol
                };

                accessDatabaseTool.UpdateRow("Contact information", "COCompanyID", $"CInt({companyId})", contactProperties);
            }
        }
    }
}
