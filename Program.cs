using Microsoft.Office.Interop.Outlook;
using System.Data;
using System.Drawing;
using System.Runtime.Versioning;
using System.Xml.Linq;

namespace OutlookToMSAccessScript
{
    [SupportedOSPlatform("windows")]
    internal class Program
    {
        static void Main(string[] args)
        {
            Print("  __  __ _                           __ _        _                            \r\n |  \\/  (_) ___ _ __ ___  ___  ___  / _| |_     / \\   ___ ___ ___  ___ ___    \r\n | |\\/| | |/ __| '__/ _ \\/ __|/ _ \\| |_| __|   / _ \\ / __/ __/ _ \\/ __/ __|   \r\n | |  | | | (__| | | (_) \\__ \\ (_) |  _| |_   / ___ \\ (_| (_|  __/\\__ \\__ \\   \r\n |_|  |_|_|\\___|_|  \\___/|___/\\___/|_|  \\__| /_/   \\_\\___\\___\\___||___/___/   \r\n     _         _                        _   _               _____           _ \r\n    / \\  _   _| |_ ___  _ __ ___   __ _| |_(_) ___  _ __   |_   _|__   ___ | |\r\n   / _ \\| | | | __/ _ \\| '_ ` _ \\ / _` | __| |/ _ \\| '_ \\    | |/ _ \\ / _ \\| |\r\n  / ___ \\ |_| | || (_) | | | | | | (_| | |_| | (_) | | | |   | | (_) | (_) | |\r\n /_/   \\_\\__,_|\\__\\___/|_| |_| |_|\\__,_|\\__|_|\\___/|_| |_|   |_|\\___/ \\___/|_|\r\n                                                                              ");
            Print("A handy dandy program created by Dez Boyle");
            PrintRainbow("------------------------------------------\n");

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

            Print("Database path: " + databasePath);
            Print("If this path is incorrect, close the program and exit the databasePath.txt file");
            Print("Select the folder in Outlook that contains the emails to import into Access\n    (you might have to click Outlook to see the prompt)\n", ConsoleColor.Green);

            //List emails
            OutlookEmailTool outlookEmailTool = new OutlookEmailTool();
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
            
            //get yes / no input from user
            while(true)
            {
                Print("Continue? (y/n)", ConsoleColor.Green);
                char key = Console.ReadKey().KeyChar;
                if (key == 'y')
                { break; }
                else if(key == 'n')
                {
                    Print("\nCancelled.  Press any key to quit", ConsoleColor.Red);
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }
            Console.WriteLine("");

            //parse email stuff here
            AccessDatabaseTool accessDatabaseTool = new AccessDatabaseTool(databasePath);
            foreach (MailItem mailItem in mailItems)
            {
                string fullName = RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "Name: ", "Email: "));
                string firstName = fullName.Split(' ')[0];
                string lastName = fullName.Split(firstName)[1].Replace(" ", string.Empty);
                string email = RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "Email: ", "<"));
                string companyName = RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "Company: ", "U.S. phone number: "));
                string phoneNumber = GetNumbersFromString(RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "U.S. phone number: ", "Address: ")));
                string address = RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "Address: ", "City: "));
                string city = RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "City: ", "State: "));
                string state = RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "State: ", "U.S. ZIP code: "));
                string zip = RemoveTrailingSpacesAndNewLines(GetTextBetween(mailItem.Body, "U.S. ZIP code: ", "reCAPTCHA:"));

                Print($"    [firstName: {firstName}]   [lastName: {lastName}]   [email: {email}]   [companyName: {companyName}]   [phoneNumber: {phoneNumber}]   [address: {address}]   [city: {city}]   [state: {state}]   [zip: {zip}]");

                //if the company doesnt exist, add it
                bool companyExists = accessDatabaseTool.RowExists("CompanyName", "CompanyName", $"'{companyName}'");
                Console.WriteLine("Company Exists: " + companyExists);
                if (!companyExists)
                {
                    Console.WriteLine("new company- Added company to 'CompanyName' table");
                    accessDatabaseTool.AddRow("CompanyName", "CompanyName", companyName);
                }

                //a list of properties to be added into the record (format: column, value)
                KeyValuePair<string, string>[] companyProperties = new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("COAddress", address),
                    new KeyValuePair<string, string>("COCity", city),
                    new KeyValuePair<string, string>("COState", state),
                    new KeyValuePair<string, string>("COZip", zip),
                };

                //update the company record
                accessDatabaseTool.UpdateRow("CompanyName", "CompanyName", $"'{companyName}'", companyProperties);
                //grab the company id (may have multiple, just take the first one)
                string companyId = accessDatabaseTool.GetRows("CompanyName", "CompanyName", $"'{companyName}'").Rows[0][0].ToString();

                //if the contact doesnt exist, add it
                bool contactExists = accessDatabaseTool.RowExists("Contact information", "COCompanyID", $"CInt({companyId})");
                if (!contactExists)
                {
                    Console.WriteLine("new contact- Added contact to 'Contact information' table");
                    accessDatabaseTool.AddRow("Contact information", "COCompanyID", companyId);
                }

                //a list of properties to be added into the record (format: column, value)
                KeyValuePair<string, string>[] contactProperties = new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("ContactFName", firstName),
                    new KeyValuePair<string, string>("ContactLName", lastName),
                    new KeyValuePair<string, string>("ContactEmail", email),
                    new KeyValuePair<string, string>("ContractPhone", phoneNumber), //they spelled Contact wrong in the database lol
                };

                //update the contact record
                accessDatabaseTool.UpdateRow("Contact information", "COCompanyID", $"CInt({companyId})", contactProperties);
                //grab the contact id (may have multiple, just take the first one)
                string contactID = accessDatabaseTool.GetRows("Contact information", "COCompanyID", $"CInt({companyId})").Rows[0][0].ToString();

                //TESTING until we can parse the bid number in the email
                // Print("TEST: enter the bid number of " + mailItem.Subject);
                // string bidNumber = Console.ReadLine();

                //Parse the Bid ID from the subject line
                string bidNumber = "NO BID ID FOUND IN EMAIL SUBJECT";
                string bidIDText = "Bid ID:"; //the text to parse for in the email subject
                int startIndex = mailItem.Subject.IndexOf(bidIDText);
                if (startIndex < 0)
                {
                    Print(bidNumber, ConsoleColor.Red);
                    continue;
                }
                startIndex += bidIDText.Length;
                bidNumber = mailItem.Subject.Substring(startIndex);
                bidNumber = bidNumber.Trim(' ');

                //a list of properties to be added into the record (format: column, value)
                KeyValuePair<string, string>[] bidProperties = new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("BidIDNo", bidNumber),
                    new KeyValuePair<string, string>("CompanyID", companyId),
                    new KeyValuePair<string, string>("ContactID", contactID)
                };

                //if the bid recipient doesnt exist, add it
                bool bidRecipientExists = accessDatabaseTool.RowExists("tbl-BidRecipients", bidProperties);
                if(!bidRecipientExists)
                {
                    Print("new bid recipient- added to 'tbl-BidRecipients' table");
                    accessDatabaseTool.AddRow("tbl-BidRecipients", bidProperties);
                }
            }

            Print("");
            Print("Finished updating database!  Press any key to close\n", ConsoleColor.Green);
            Print(" /\\_/\\ ♥\r\n >^,^<\r\n  / \\\r\n (___)_/");
            Console.ReadKey();
            Environment.Exit(0);
        }
        private static string GetNumbersFromString(string input)
        {
            return new string(input.Where(c => char.IsDigit(c)).ToArray());
        }

        private static string GetTextBetween(string source, string start, string end)
        {
            int pFrom = source.IndexOf(start) + start.Length;
            int pTo = source.LastIndexOf(end);

            return source.Substring(pFrom, pTo - pFrom);
        }

        private static string RemoveTrailingSpacesAndNewLines(string text)
        {
            text = text.Trim(' ');
            text = text.Replace("\n", string.Empty);
            text = text.Replace("\r", string.Empty);
            return text;
        }

        private static void Print(string text, ConsoleColor color = ConsoleColor.White)
        {
            Console.ForegroundColor = color;    
            Console.WriteLine(text);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private static void PrintRainbow(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                ConsoleColor color = ConsoleColor.White;
                switch (i % 6)
                {
                    case 0: color = ConsoleColor.Red; break;
                    case 1: color = ConsoleColor.White; break;
                    case 2: color = ConsoleColor.Green; break;
                    case 3: color = ConsoleColor.Cyan; break;
                    case 4: color = ConsoleColor.Blue; break;
                    case 5: color = ConsoleColor.Magenta; break;
                }
                Console.ForegroundColor = color;
                Console.Write(text[i]);
                Console.ForegroundColor = ConsoleColor.White;
            }
            Console.WriteLine();
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
