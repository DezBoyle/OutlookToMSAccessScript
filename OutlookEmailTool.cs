using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;


namespace OutlookToMSAccessScript
{
    internal class OutlookEmailTool
    {
        public Microsoft.Office.Interop.Outlook.Items GetEmails()
        {
            try
            {
                // Create the Microsoft.Office.Interop.Outlook application.
                // in-line initialization
                Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();

                // Get the MAPI namespace.
                Microsoft.Office.Interop.Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                // Log on by using the default profile or existing session (no dialog box).
                oNS.Logon(Missing.Value, Missing.Value, false, true);

                // Alternate logon method that uses a specific profile name.
                // TODO: If you use this logon method, specify the correct profile name
                // and comment the previous Logon line.
                //oNS.Logon("profilename",Missing.Value,false,true);

                //Get the Inbox folder.
                //Microsoft.Office.Interop.Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                Microsoft.Office.Interop.Outlook.MAPIFolder oInbox = oNS.PickFolder();

                //Get the Items collection in the Inbox folder.
                Microsoft.Office.Interop.Outlook.Items oItems = oInbox.Items;

                // Get the first message.
                // Because the Items folder may contain different item types,
                // use explicit typecasting with the assignment.
                //Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oItems.GetFirst();

                //Output some common properties.
                //Console.WriteLine(oMsg.Subject);
                //Console.WriteLine(oMsg.SenderName);
                //Console.WriteLine(oMsg.ReceivedTime);
                //Console.WriteLine(oMsg.Body);

                //Check for attachments.
                //int AttachCnt = oMsg.Attachments.Count;
                //Console.WriteLine("Attachments: " + AttachCnt.ToString());

                //TO DO: If you use the Microsoft Microsoft.Office.Interop.Outlook 10.0 Object Library, uncomment the following lines.
                /*if (AttachCnt > 0) 
                {
                for (int i = 1; i <= AttachCnt; i++) 
                 Console.WriteLine(i.ToString() + "-" + oMsg.Attachments.Item(i).DisplayName);
                }*/

                //TO DO: If you use the Microsoft Microsoft.Office.Interop.Outlook 11.0 Object Library, uncomment the following lines.
                /*if (AttachCnt > 0) 
                {
                for (int i = 1; i <= AttachCnt; i++) 
                 Console.WriteLine(i.ToString() + "-" + oMsg.Attachments[i].DisplayName);
                }*/

                //Display the message.
                //oMsg.Display(true); //modal

                //Log off.
                oNS.Logoff();

                //Explicitly release objects.
                //oMsg = null;
                oInbox = null;
                oNS = null;
                oApp = null;

                return oItems;
            }

            //Error handler.
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("{0} Exception caught: ", e);
                Console.ForegroundColor = ConsoleColor.White;
                return null;
            }
        }
    }
}
