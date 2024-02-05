using Microsoft.Office.Interop.Outlook;
using System.Runtime.Versioning;

namespace OutlookToMSAccessScript
{
    [SupportedOSPlatform("windows")]
    internal class Program
    {
        static void Main(string[] args)
        {
            Print("  __  __ _                           __ _        _                            \r\n |  \\/  (_) ___ _ __ ___  ___  ___  / _| |_     / \\   ___ ___ ___  ___ ___    \r\n | |\\/| | |/ __| '__/ _ \\/ __|/ _ \\| |_| __|   / _ \\ / __/ __/ _ \\/ __/ __|   \r\n | |  | | | (__| | | (_) \\__ \\ (_) |  _| |_   / ___ \\ (_| (_|  __/\\__ \\__ \\   \r\n |_|  |_|_|\\___|_|  \\___/|___/\\___/|_|  \\__| /_/   \\_\\___\\___\\___||___/___/   \r\n     _         _                        _   _               _____           _ \r\n    / \\  _   _| |_ ___  _ __ ___   __ _| |_(_) ___  _ __   |_   _|__   ___ | |\r\n   / _ \\| | | | __/ _ \\| '_ ` _ \\ / _` | __| |/ _ \\| '_ \\    | |/ _ \\ / _ \\| |\r\n  / ___ \\ |_| | || (_) | | | | | | (_| | |_| | (_) | | | |   | | (_) | (_) | |\r\n /_/   \\_\\__,_|\\__\\___/|_| |_| |_|\\__,_|\\__|_|\\___/|_| |_|   |_|\\___/ \\___/|_|\r\n                                                                              ");
            Print("A handy dandy program created by Dez Boyle", ConsoleColor.White);

            string saveFilePath = "databasePath.txt";
            string databasePath = "C:/Users/dboyle/testdb.mdb";

            if(File.Exists(saveFilePath))
            { databasePath = File.ReadAllText(saveFilePath); }
            else
            {
                Print("Enter the path of the database file (.mdb)", ConsoleColor.Yellow);
                File.WriteAllText(saveFilePath, Console.ReadLine());
            }

            File.WriteAllText("information.txt", "Created by Dez Boyle\nSource Code: https://github.com/DezBoyle/OutlookToMSAccessScript");

            //DebugPrompt(databasePath);

            OutlookEmailTool outlookEmailTool = new OutlookEmailTool();

            Print("Select the folder that contains the emails to import into Access", ConsoleColor.Green);

            Items emails = outlookEmailTool.GetEmails();

            List<MailItem> mailItems = new List<MailItem>();
            for (int i = 1; i < emails.Count + 1; i++)
            {
                if (emails[i] as MailItem != null)
                { mailItems.Add(emails[i]); }
            }

            for (int i = 0; i < mailItems.Count; i++)
            {
                MailItem mailItem = mailItems[i];
                Print("    > " + mailItem.Subject, i % 2 == 0 ? ConsoleColor.White : ConsoleColor.Gray);
            }

            Print("\nThe above^ emails will be entered into Access.", ConsoleColor.Green);
            
            while(true)
            {
                Print("Continue? (y/n)", ConsoleColor.Green);
                char key = Console.ReadKey().KeyChar;
                if (key == 'y')
                { break; }
                else if(key == 'n')
                {
                    Print("Cancelled.  Press any key to quit", ConsoleColor.Red);
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }

            //parse email stuff here

        }

        private static void Print(string text, ConsoleColor color = ConsoleColor.White)
        {
            Console.ForegroundColor = color;    
            Console.WriteLine(text);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private static void DebugPrompt(string databasePath)
        {
            AccessDatabaseTool accessDatabaseTool = new AccessDatabaseTool(databasePath);

            while (true)
            {
                Console.WriteLine("(TEST) Enter Company Name:");
                string companyName = Console.ReadLine();

                bool companyExists = accessDatabaseTool.RowExists("CompanyName", "CompanyName", $"'{companyName}'");
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

                string companyId = accessDatabaseTool.GetRows("CompanyName", "CompanyName", $"'{companyName}'").Rows[0][0].ToString(); //grab the company id (may have multiple, just take the first one)

                bool contactExists = accessDatabaseTool.RowExists("Contact information", "COCompanyID", $"CInt({companyId})");
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
