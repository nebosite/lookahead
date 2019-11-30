using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;

namespace LookAhead
{
    class Program
    {
        static Outlook.Application oOutlook;
        static Outlook.NameSpace oNameSpace;
        static Outlook.OlExchangeConnectionMode connectionMode;

        static void Main(string[] args)
        {
            var toAddress = (string)null;
            if (args.Length > 0) toAddress = args[0];
            try
            {
                oOutlook = OutlookHelper.GetApplicationObject();
                Outlook.Folders folders = oOutlook.Session.Folders;
                List<HeadsUpItem> returnItems = new List<HeadsUpItem>();
                ProcessFolder((Outlook.Folder)oOutlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar), returnItems);
                ProcessFolder((Outlook.Folder)oOutlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks), returnItems);

                StringBuilder htmlText = new StringBuilder();
                htmlText.Append("<HEAD>" +
                    "<STYLE TYPE=\"text/css\">" +
                    "<!--" +
                    "BODY" +
                    "   {" +
                    "       font-family:sans-serif;" +
                    "       font-size: 10px;" +
                    "   }" +
                    "B" +
                    "   {" +
                    "       font-family:sans-serif;" +
                    "       font-size: 10px;" +
                    "       font-style: Bold;" +
                    "   }" +
                    "-->" +
                    "</STYLE>" +
                    "</HEAD>");

                for (int i = 0; i < 7; i++)
                {
                    AppendHeadsUp(htmlText, returnItems, DateTime.Now.AddDays(i).Date);
                }

                Outlook.MailItem newMail = oOutlook.CreateItem(Outlook.OlItemType.olMailItem);
                if(toAddress == null)
                {
                    toAddress = GetCurrentUserEmailAddress();
                }
                newMail.To = toAddress;
                newMail.Subject = "Heads Up!";
                newMail.HTMLBody = htmlText.ToString();
                //newMail.SaveSentMessageFolder = (Outlook.Folder)oOutlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                newMail.Send();
            }
            catch (ApplicationException e)
            {
                Console.WriteLine("ERROR: " + e.Message);
            }

           // Console.Read();
        }

        private static string GetCurrentUserEmailAddress()
        {
            try
            {
                string currentUserAddress = Environment.UserName;
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(oOutlook.Session.AutoDiscoverXml);
                XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
                namespaceManager.AddNamespace("ns1", "http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006");
                namespaceManager.AddNamespace("ns2", "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a");

                XmlNode addressNode = xmlDoc.SelectSingleNode("//ns1:Autodiscover/ns2:Response/ns2:User/ns2:AutoDiscoverSMTPAddress", namespaceManager);
                if (addressNode != null) currentUserAddress = addressNode.InnerText;
                return currentUserAddress;
            }
            catch(Exception e)
            {
                throw new ApplicationException("Error trying to automatically obtain user email address: " + e.Message, e);
            }
        }

        private static void AppendHeadsUp(StringBuilder htmlText, List<HeadsUpItem> returnItems, DateTime start)
        {
            DateTime end = start.AddHours(24);
            htmlText.Append("<B>" + start.ToLongDateString() + "</B>");
            var items = from item in returnItems
                        where item.Start >= start && item.Start < end
                        select item;
            if (items.Count<HeadsUpItem>() == 0)
            {
                htmlText.Append("<TABLE><TD>*** No items for this day ***</TD></TABLE>");
            }
            else
            {
                htmlText.Append("<TABLE>");
                foreach (HeadsUpItem item in items)
                {
                    string fontStart = "";
                    string fontEnd = "";
                    string bgColor = "bgColor=Yellow";
                    if (item.Ignorable)
                    {
                        bgColor = "";
                        fontStart = "<FONT color='#aaaaaa'>";
                        fontEnd = "</FONT>";
                    }

                    htmlText.Append("<TR " + bgColor + ">");
                    htmlText.Append("<TD width=20></TD>");
                    htmlText.Append("<TD halign=right>" + fontStart + item.Start.ToLongTimeString() + " (" + (item.End - item.Start).TotalMinutes.ToString("0") + "m) " + fontEnd + "</TD>");
                    htmlText.Append("<TD width=20></TD>");
                    htmlText.Append("<TD>" + fontStart + item.Title + " (" + item.Location + fontEnd + ")</TD>");
                    htmlText.Append("</TR>");

                }
                htmlText.Append("</TABLE>");

            }
        }

        private static void ProcessFolder(Outlook.Folder folder, List<HeadsUpItem> returnItems)
        {
            
            Console.WriteLine(folder.Name);

            DateTime now = DateTime.Now;
            // Set end value
            DateTime lookAhead = now.AddDays(7);
            // Initial restriction is Jet query for date range
            string timeSlotFilter = 
                "[Start] >= '" + now.ToString("g")
                + "' AND [Start] <= '" + lookAhead.ToString("g") + "'";

            Outlook.Items items = folder.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            items = items.Restrict(timeSlotFilter);
            foreach (object item in items)
            {
                Outlook.AppointmentItem appointment = item as Outlook.AppointmentItem;
                Outlook.TaskItem task = item as Outlook.TaskItem;

                if (appointment != null)
                {
                    HeadsUpItem newItem = new HeadsUpItem();
                    newItem.Title = appointment.Subject;
                    newItem.Start = appointment.Start;
                    newItem.End = appointment.End;
                    newItem.Location = appointment.Location;
                    newItem.Ignorable = true;
                    Outlook.RecurrencePattern pattern = appointment.GetRecurrencePattern();
                    if (appointment.RecurrenceState == Outlook.OlRecurrenceState.olApptOccurrence && 
                        pattern.RecurrenceType == Outlook.OlRecurrenceType.olRecursWeekly)
                    {
                        Console.WriteLine("  Ignore:" + appointment.Subject);
                    }
                    else
                    {
                        newItem.Ignorable = false;
                        Console.WriteLine("  A: " + appointment.Subject);
                        Console.WriteLine("      S: " + appointment.Start.ToString("g"));
                        Console.WriteLine("      E: " + appointment.End.ToString("g"));
                    }

                    returnItems.Add(newItem);


                }
                else if (task != null)
                {
                    Console.WriteLine("  T: " + task.Subject);
                    
                }
                else
                {
                    Console.WriteLine("  ERR: Unknown Type: " + item.GetType());
                }
            } 


            foreach (Outlook.Folder subFolder in folder.Folders)
            {
                ProcessFolder(subFolder, returnItems);
            }

        }

    }

    public class HeadsUpItem
    {
        public string Title { get; set; }
        public string Location { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public bool Ignorable { get; set; }
    }
}
