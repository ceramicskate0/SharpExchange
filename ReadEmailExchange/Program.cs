using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System.IO;

namespace ReadEmailExchange
{
    public class Program
    {
        public static int _pageSize = 200;
        public Folder _imHistoryFolder;
        public static string OutputResultsFilename = "ExchangeDump";
        public List<EmailMessage> _imHistory;
        public static ExchangeService es = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
        public static LyncConversationHistory lyncConvHistory = new LyncConversationHistory();
        public static List<OutlookData> EmailInbox = new List<OutlookData>();
        public static List<OutlookData> EmailSent = new List<OutlookData>();
        public static List<OutlookData> EmailCalendar = new List<OutlookData>();
        public static List<OutlookData> EmailSkype = new List<OutlookData>();
        public static List<OutlookData> EmailDrafts = new List<OutlookData>();
        public static List<OutlookData> EmailDeleted = new List<OutlookData>();
        public static List<string> EmailFolder = new List<string>();

        public static string GetPlainTextFromHtml(string htmlString)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlString);
            return doc.DocumentNode.FirstChild.InnerText+" ";
        }

        /// <summary>
        /// 0=Emails
        /// 1=Skype
        /// </summary>
        /// <param name="data"></param>
        /// <param name="Username"></param>
        /// <param name="Dump"></param>
        /// <param name="OutputType"></param>
        public static void OUTtoCSV(List<OutlookData> data, string Username,string Dump,int OutputType=0)
        {
            string filename = Username + "_" +Dump +"_"+ OutputResultsFilename + ".csv";

            Console.WriteLine("[*] Writing output file to :" + filename);

            if (File.Exists(filename) == false)
            {
                File.WriteAllText(filename, "\nInboxFolders: " + string.Join("| ", EmailFolder.ToArray()) + "\n");
                File.AppendAllText(filename, "To" + "," + "From" + "," + "Subject" + "," + "CC" + "," + "Attachment_Count"+","+ "Body" + "\n");
            }
            if (OutputType == 0)//Emails
            {
                for (int x = 0; x < data.Count; ++x)
                {
                    File.AppendAllText(filename, data.ElementAt(x).to.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).from.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).subj.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).cc.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).AttachCount.ToString() + "," + data.ElementAt(x).body.Replace("\r", " ").Replace("\n", "").Replace(",", "").Replace("&nbsp;", "") + "," + "\n");
                }
            }
            else if (OutputType==1)//Skype
            {
                for (int x = 0; x < data.Count; ++x)
                {
                    File.AppendAllText(filename, data.ElementAt(x).to.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).from.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).subj.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).cc.Replace("\n", " ").Replace("\r", " ").Replace(",", "") + "," + data.ElementAt(x).AttachCount.ToString() + "," + data.ElementAt(x).body.Replace("\r", "").Replace(",", "").Replace("&nbsp;", "") + "," + "\n------------------------------------------------\n");
                }
            }
        }

        public static void DumpSkype()
        {
            try
            {
                lyncConvHistory.QueryImHistory(new string[] { " ","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","1","2","3","4","5","6","7","8","9","0" }, new DateTime(DateTime.Now.Year - 1, 1, 1), es, _pageSize);

                IEnumerable<Item> items = lyncConvHistory.RetrieveSpecialEmailFolderData(es, _pageSize);
                Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " Skype/Lync conversation history");

                if (items != null)
                {
                    foreach (Item item in items)
                    {
                        OutlookData obj = new OutlookData();
                        item.Load();
                        obj.to = item.DisplayTo;
                        obj.from = item.DisplayTo;
                        obj.to = item.DisplayTo;
                        obj.subj = item.Subject;
                        obj.cc = item.DisplayTo;
                        obj.body = GetPlainTextFromHtml(item.Body.Text);
                        EmailSkype.Add(obj);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[-] DumpSkype ERROR: " + e.Message.ToString());
            }
        }

        public static void DumpInbox()
        {
            try 
            { 
            FindItemsResults<Item> Inbox = es.FindItems(WellKnownFolderName.Inbox, new ItemView(_pageSize));
            Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " inbox emails");
                foreach (Item item in Inbox)
            {
                OutlookData obj = new OutlookData();
                EmailMessage msg = EmailMessage.Bind(es, item.Id);

                try
                {
                    if (msg.Body.BodyType.ToString().ToLower().Contains("html"))
                    {
                        obj.body = GetPlainTextFromHtml(msg.Body.Text+" ");
                    }
                    else
                    {
                        obj.body = msg.Body.Text+" ";
                    }
                    obj.AttachCount = msg.Attachments.Count+0;
                    obj.to = msg.ReceivedBy.Address + " ";
                    obj.cc = msg.DisplayCc+ " ";
                    obj.from = msg.From.Address + " ";
                    obj.subj = msg.Subject + " ";
                    EmailInbox.Add(obj);
                }
                catch (Exception e)
                {
                        Console.WriteLine(" [-] Record ERROR:" + e.Message.ToString());
                }
            }
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] DumpInbox ERROR: " + e.Message.ToString());
            }
            DumpInboxFolders();
        }

        public static void DumpSentItems()
        {
            try
            {
                FindItemsResults<Item> SentItems = es.FindItems(WellKnownFolderName.SentItems, new ItemView(_pageSize));
                Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " SentItems");

                foreach (Item item in SentItems)
                {
                    OutlookData obj = new OutlookData();
                    EmailMessage msg = EmailMessage.Bind(es, item.Id);

                    try
                    {
                        obj.AttachCount = msg.Attachments.Count;
                        if (msg.Body.BodyType.ToString().ToLower().Contains("html"))
                        {
                            obj.body = GetPlainTextFromHtml(msg.Body.Text);
                        }
                        else
                        {
                            obj.body = msg.Body.Text;
                        }
                        obj.to = msg.DisplayTo + " ";
                        obj.cc = " ";
                        obj.from = msg.From.Address + " ";
                        obj.subj = msg.Subject+" ";
                        EmailSent.Add(obj);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" [-] Record ERROR:" + e.Message.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] DumpSentItems ERROR: " + e.Message.ToString());
            }
        }

        public static void DumpDrafts()
        {
            try
            {
                FindItemsResults<Item> Drafts = es.FindItems(WellKnownFolderName.Drafts, new ItemView(_pageSize));
                Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " Drafts");

                foreach (Item item in Drafts)
                {
                    OutlookData obj = new OutlookData();
                    EmailMessage msg = EmailMessage.Bind(es, item.Id);

                    try
                    {
                        obj.AttachCount = msg.Attachments.Count;
                        if (msg.Body.BodyType.ToString().ToLower().Contains("html"))
                        {
                            obj.body = GetPlainTextFromHtml(msg.Body.Text);
                        }
                        else
                        {
                            obj.body = msg.Body.Text;
                        }
                        obj.to = msg.DisplayTo + " ";
                        obj.cc = msg.DisplayCc + " ";
                        obj.from = msg.From.Address + " ";
                        obj.subj = msg.Subject + " ";
                        EmailDrafts.Add(obj);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" [-] Record ERROR:" + e.Message.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] DumpDrafts ERROR: " + e.Message.ToString());
            }
        }

        public static void DumpDeleted()
        {
            try
            {
                FindItemsResults<Item> DeletedItems = es.FindItems(WellKnownFolderName.DeletedItems, new ItemView(_pageSize));
                Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " DeletedItems");

                foreach (Item item in DeletedItems)
                {
                    OutlookData obj = new OutlookData();
                    EmailMessage msg = EmailMessage.Bind(es, item.Id);

                    try
                    {
                        obj.AttachCount = msg.Attachments.Count;
                        if (msg.Body.BodyType.ToString().ToLower().Contains("html"))
                        {
                            obj.body = GetPlainTextFromHtml(msg.Body.Text);
                        }
                        else
                        {
                            obj.body = msg.Body.Text;
                        }
                        obj.to = msg.DisplayTo + " ";
                        obj.cc = msg.DisplayCc + " ";
                        obj.from = msg.From.Address + " ";
                        obj.subj = msg.Subject + " ";
                        EmailDeleted.Add(obj);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" [-] Record ERROR:"+e.Message.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] DumpDeleted ERROR: " + e.Message.ToString());
            }
        }

        public static void DumpInboxFolders()
        {
            Console.WriteLine("[*] Trying to get Inbox folders");
            lyncConvHistory.QueryImHistory(new string[] { " " }, new DateTime(DateTime.Now.Year - 1, 1, 1), es, 50);
            IEnumerable<Item> items = lyncConvHistory.RetrieveSpecialEmailFolderData(es, 50, "conversation history",true); 
            EmailFolder = EmailFolder.Distinct().ToList();
        }

        public static void DumpCalendarItems()
        {
            try
            {
                FindItemsResults<Item> CalendarItems = es.FindItems(WellKnownFolderName.Calendar, new ItemView(_pageSize));
                Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " Calendar Items");

                foreach (Item item in CalendarItems)
                {
                    OutlookData obj = new OutlookData();
                    EmailMessage msg = EmailMessage.Bind(es, item.Id);

                    try
                    {
                        obj.AttachCount = msg.Attachments.Count;
                        obj.body = msg.Body.Text;
                        obj.to = msg.DisplayTo;
                        obj.cc = msg.DisplayCc;
                        obj.from = msg.From.Address;
                        obj.subj = msg.Subject;
                        EmailSent.Add(obj);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" [-] Record ERROR:" + e.Message.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] DumpCalendarItems ERROR: " + e.Message.ToString());
            }
        }

        public static void HelpMenu()
        {
            Console.WriteLine(@"

            Required Inputs (Must be in order shown):

            ReadEmailExchange.exe WEBDomain DomainName Password InternalDomainName DUMPItem
                Example WEBDomain: webmail.domain.com
                Example DomainName: User1
                Example Password: SecretPassword
                Example InternalDomainName: domain
                
                Options for DUMPItem:
                    Inbox
                    Sent
                    Drafts
                    Deleted
                    Skype
                    Attachments (Will Download Atatchments from the Inbox, DeletedItems, and Sent Items folders)
                    SendEmail ToEmailAddress~Subject~Body(Body can be file path)~AttachmentLocalFilePath(optional)
                    All (All == will try to dump all the items above)(I would default to this if unsure)

            Optional Inputs:

            ReadEmailExchange.exe WEBDomain DomainName Password InternalDomainName DUMPItem NumberOfSearchResultsToReturn
                            Example NumberOfSearchResultsToReturn (will return a maximum of the number,default 10): 10
                            Note: NumberOfSearchResultsToReturn must be a int/whole number

            Optional Inputs:

            ReadEmailExchange.exe WEBDomain DomainName Password InternalDomainName DUMPItem NumberOfSearchResultsToReturn OutputFileNameOrPath
                Example OutputFileNameOrPath: C:\file.csv
                Note: Program needs permission to write to location
        ");
            Environment.Exit(1);
        }

        public static void GetAttachmentsFromEmail(string Username)
        {
            FindItemsResults<Item> InboxFolder = es.FindItems(WellKnownFolderName.Inbox, new ItemView(_pageSize));
            FindItemsResults<Item> SentFolder = es.FindItems(WellKnownFolderName.SentItems, new ItemView(_pageSize));
            FindItemsResults<Item> DeletedFolder = es.FindItems(WellKnownFolderName.DeletedItems, new ItemView(_pageSize));

            Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " Inbox Attachments Items");

            foreach (Item item in InboxFolder)
            {
                EmailMessage message = EmailMessage.Bind(es, item.Id, new PropertySet(ItemSchema.Attachments));

                // Iterate through the attachments collection and load each attachment.
                foreach (Attachment attachment in message.Attachments)
                {
                    try
                    {
                        if (attachment is FileAttachment)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            // Load the attachment into a file.
                            // This call results in a GetAttachment call to EWS.
                            if (fileAttachment.Name.Contains(".png") == false && fileAttachment.Name.Contains(".gif") == false && fileAttachment.Name.Contains(".txt") == false && fileAttachment.Name.Contains(".htm") == false && fileAttachment.Name.Contains(".jpg") == false)
                            {
                                fileAttachment.Load(Username + "_Inbox_" + fileAttachment.Name);
                                Console.WriteLine(" [+] File attachment name: " + fileAttachment.Name);

                                // Write the bytes of the attachment into a file.
                            }
                        }
                        else // Attachment is an item attachment.
                        {
                            ItemAttachment itemAttachment = attachment as ItemAttachment;
                            // Load attachment into memory and write out the subject.
                            // This does not save the file like it does with a file attachment.
                            // This call results in a GetAttachment call to EWS.
                            itemAttachment.Load();
                            Console.WriteLine("Item attachment name: " + itemAttachment.Name);
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" [-] Record ERROR:" + e.Message.ToString());
                    }
                }
            }
            
            Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " Sent Items Attachments Items");       
            foreach (Item item in SentFolder)
            {
                EmailMessage message = EmailMessage.Bind(es, item.Id, new PropertySet(ItemSchema.Attachments));

                // Iterate through the attachments collection and load each attachment.
                foreach (Attachment attachment in message.Attachments)
                {
                    try
                    {
                        if (attachment is FileAttachment)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            // Load the attachment into a file.
                            // This call results in a GetAttachment call to EWS.
                            if (fileAttachment.Name.Contains(".png") == false && fileAttachment.Name.Contains(".gif") == false && fileAttachment.Name.Contains(".txt") == false && fileAttachment.Name.Contains(".htm") == false && fileAttachment.Name.Contains(".jpg") == false)
                            {
                                fileAttachment.Load(Username + "_Sent_" + fileAttachment.Name);
                                Console.WriteLine(" [+] File attachment name: " + fileAttachment.Name);

                                // Write the bytes of the attachment into a file.
                            }
                        }
                        else // Attachment is an item attachment.
                        {
                            ItemAttachment itemAttachment = attachment as ItemAttachment;
                            // Load attachment into memory and write out the subject.
                            // This does not save the file like it does with a file attachment.
                            // This call results in a GetAttachment call to EWS.
                            itemAttachment.Load();
                            Console.WriteLine("Item attachment name: " + itemAttachment.Name);
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" [-] Record ERROR:" + e.Message.ToString());
                    }
                }
            }
            
            Console.WriteLine("[*] Retrieving the first " + _pageSize.ToString() + " Deleted Items Attachments Items");

            foreach (Item item in DeletedFolder)
            {
                EmailMessage message = EmailMessage.Bind(es, item.Id, new PropertySet(ItemSchema.Attachments));

                // Iterate through the attachments collection and load each attachment.
                foreach (Attachment attachment in message.Attachments)
                {
                    try
                    {
                        if (attachment is FileAttachment)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            // Load the attachment into a file.
                            // This call results in a GetAttachment call to EWS.
                            if (fileAttachment.Name.Contains(".png") == false && fileAttachment.Name.Contains(".gif") == false && fileAttachment.Name.Contains(".txt") == false && fileAttachment.Name.Contains(".htm") == false && fileAttachment.Name.Contains(".jpg") == false)
                            {
                                fileAttachment.Load(Username + "_Deleted_" + fileAttachment.Name);
                                Console.WriteLine(" [+] File attachment name: " + fileAttachment.Name);

                                // Write the bytes of the attachment into a file.
                            }
                        }
                        else // Attachment is an item attachment.
                        {
                            ItemAttachment itemAttachment = attachment as ItemAttachment;
                            // Load attachment into memory and write out the subject.
                            // This does not save the file like it does with a file attachment.
                            // This call results in a GetAttachment call to EWS.
                            itemAttachment.Load();
                            Console.WriteLine("Item attachment name: " + itemAttachment.Name);
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" [-] Record ERROR:" + e.Message.ToString());
                    }
                }
            }
        }

        public static void SendEmail(string TOEmail,string subj, string Body, string filepath="")
        {
            // Create a new email message. 
            EmailMessage message = new EmailMessage(es);
            // Specify the email recipient and subject. 
            message.ToRecipients.Add(TOEmail);
            message.Subject = subj;
            // Identify the extended property that can be used to specify when to send the email. 
            ExtendedPropertyDefinition PidTagDeferredSendTime = new ExtendedPropertyDefinition(16367, MapiPropertyType.SystemTime);
            // Set the time that will be used to specify when the email is sent. 
            // In this example, the email will be sent one minute after the next line executes, 
            // provided that the message.SendAndSaveCopy request is processed by the server within one minute. 
            string sendTime = DateTime.Now.AddSeconds(10).ToUniversalTime().ToString();
            // Specify when to send the email by setting the value of the extended property. 
            message.SetExtendedProperty(PidTagDeferredSendTime, sendTime);
            
            // Specify the email body. 
            if (File.Exists(Body))
            {
                try
                {
                    message.Body = File.ReadAllText(Body);
                }
                catch (Exception e)
                {
                    Console.WriteLine(" [!] Body File error makeing body be blank error was: " + e.Message.ToString());
                    message.Body = "";
                }
            }
            else
            {
                message.Body = Body;
            }

            if (string.IsNullOrEmpty(filepath)==false)
            {
                message.Attachments.AddFileAttachment(filepath);
            }
            message.Send();
        }
        
        public static bool IsString(object value)
        {
            return value is string;
        }

        static void Main(string[] args)
        {
            try
            {
                string URL = args[0];//ie webmail.DOMAIN.com
                string DomainUserName = args[1];//Username on the Domain not email
                string Password = args[2];//that account password
                string domainName = args[3];//Internal Domain Name. The NMAP NTLM u get fronm /ews/
                string dump = args[4];
                if (args.Length > 5 && IsString(args[5])==false)
                {
                    _pageSize = Convert.ToInt32(args[5]);//Max number of results to try to return
                }
                if (args.Length > 6)
                {
                    OutputResultsFilename = args[6];
                }
                AutodiscoverService autodiscoverService = new AutodiscoverService(URL);
                es.Credentials = new WebCredentials(DomainUserName, Password, domainName);
                es.UseDefaultCredentials = false;
                Uri redirectionUri = new Uri(@"https://" + URL + "/EWS/Exchange.asmx");
                es.Url = redirectionUri;
                Console.WriteLine("[*] Trying to connect to Exchange EWS Service...");
                if (es != null)
                {
                    switch (dump.ToLower())
                    {
                        case "all":
                            bool badlogin = false;
                            try
                            {
                                DumpInbox();
                                OUTtoCSV(EmailInbox, DomainUserName, "Inbox");
                            }
                            catch(Exception e)
                            {
                                if (e.Message.ToString().ToLower().Contains("401"))
                                {
                                    Console.WriteLine("[!] BAD LOGIN, Will skip further attempts");
                                    badlogin = true;
                                }
                                Console.WriteLine(" [-] DumpInbox() Failed " + e.Message.ToString());
                            }
                            try
                            {
                                if (badlogin == false)
                                {
                                    DumpSentItems();
                                    OUTtoCSV(EmailSent, DomainUserName, "Sent");
                                }
                            }
                            catch (Exception e)
                            {
                                if (e.Message.ToString().ToLower().Contains("401"))
                                {
                                    Console.WriteLine("[!] BAD LOGIN, Will skip further attempts");
                                    badlogin = true;
                                }
                                Console.WriteLine(" [-] DumpSentItems() Failed " + e.Message.ToString());
                            }
                            try
                            {
                                if (badlogin == false)
                                {
                                    DumpDrafts();
                                    OUTtoCSV(EmailDrafts, DomainUserName, "Drafts");
                                }
                            }
                            catch (Exception e)
                            {
                                if (e.Message.ToString().ToLower().Contains("401"))
                                {
                                    Console.WriteLine("[!] BAD LOGIN, Will skip further attempts");
                                    badlogin = true;
                                }
                                Console.WriteLine(" [-] DumpDrafts() Failed " + e.Message.ToString());
                            }
                            try
                            {
                                if (badlogin == false)
                                {
                                    DumpDeleted();
                                    OUTtoCSV(EmailDeleted, DomainUserName, "Deleted");
                                }
                            }
                            catch (Exception e)
                            {
                                if (e.Message.ToString().ToLower().Contains("401"))
                                {
                                    Console.WriteLine("[!] BAD LOGIN, Will skip further attempts");
                                    badlogin = true;
                                }
                                Console.WriteLine(" [-] DumpDeleted() Failed " + e.Message.ToString());
                            }
                            try
                            {
                                if (badlogin == false)
                                {
                                    DumpSkype();
                                    OUTtoCSV(EmailSkype, DomainUserName, "Skype", 1);
                                }
                            }
                            catch (Exception e)
                            {
                                if (e.Message.ToString().ToLower().Contains("401"))
                                {
                                    Console.WriteLine("[!] BAD LOGIN, Will skip further attempts");
                                    badlogin = true;
                                }
                                Console.WriteLine(" [-] DumpSkype() Failed " + e.Message.ToString());
                            }
                            try
                            {
                                if (badlogin == false)
                                {
                                    GetAttachmentsFromEmail(DomainUserName);
                                }
                            }
                            catch (Exception e)
                            {
                                if (e.Message.ToString().ToLower().Contains("401"))
                                {
                                    Console.WriteLine("[!] BAD LOGIN, Will skip further attempts");
                                    badlogin = true;
                                }
                                Console.WriteLine(" [-] GetAttachmentsFromEmail() Failed " + e.Message.ToString());
                            }
                            break;
                        case "inbox":
                            DumpInbox();
                            OUTtoCSV(EmailInbox, DomainUserName,"Inbox");
                            break;
                        case "sent":
                            DumpSentItems();
                            OUTtoCSV(EmailSent, DomainUserName,"Sent");
                            break;
                        case "drafts":
                            DumpDrafts();
                            OUTtoCSV(EmailDrafts, DomainUserName,"Drafts");
                            break;
                        case "deleted":
                            DumpDeleted();
                            OUTtoCSV(EmailDeleted, DomainUserName,"Deleted");
                            break;
                        case "attachments":
                            GetAttachmentsFromEmail(DomainUserName);
                            break;
                        case "sendemail":
                            if (args.Length == 6)
                            {
                                List<string> EmailSection = args[5].Split('~').ToList();

                                if (EmailSection.Count == 3)
                                {
                                    SendEmail(EmailSection.ElementAt(0), EmailSection.ElementAt(1), EmailSection.ElementAt(2));
                                }
                                else if (EmailSection.Count == 4)
                                {
                                    SendEmail(EmailSection.ElementAt(0), EmailSection.ElementAt(1), EmailSection.ElementAt(2), EmailSection.ElementAt(3));
                                }
                                else
                                {
                                    Console.WriteLine("[!] Error wrong number of email section args.");
                                }
                            }
                            else
                            {
                                Console.WriteLine("[!] Error wrong number of input args.");
                            }
                            break;
                        case "skype":
                            DumpSkype();
                            OUTtoCSV(EmailSkype, DomainUserName,"Skype",1);
                            break;            
                    }
                    //DumpCalendarItems();
                    //OUTtoCSV(EmailCalendar);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] MAIN Error: " + e.Message.ToString());

                HelpMenu();
            }
        }
    }
    public class OutlookData
    {
        string Body = " ";
        string To = " ";
        string From = " ";
        string CC = " ";
        string SUBJ = " ";
        public string body
        {
            get { return Body; }   // get method
            set { Body = value; }  // set method
        }
        public string to
        {
            get { return To; }   // get method
            set { To = value; }  // set method
        }
        public string from
        {
            get { return From; }   // get method
            set { From = value; }  // set method
        }
        public string cc
        {
            get { return CC; }   // get method
            set { CC = value; }  // set method
        }
        public string subj
        {
            get { return SUBJ; }   // get method
            set { SUBJ = value; }  // set method
        }
        public int AttachCount
        { get; set; }
      
    }

    public class LyncConversationHistory
    {
        public Folder _imHistoryFolder = null;
        public List<EmailMessage> _imHistory = null;
        public List<EmailMessage> EmailSkypeConvo = new List<EmailMessage>();

        public IEnumerable<Item> RetrieveSpecialEmailFolderData(ExchangeService es,int _pageSize, string FolderName= "conversation history", bool GetAllfolders = false)
        {

            // Get the "Conversation History" folder, if not already found.
            if (_imHistoryFolder == null)
            {
                _imHistoryFolder = this.FindImHistoryFolder(es, _pageSize, FolderName, GetAllfolders);
                if (_imHistoryFolder == null)
                {
                    return null;
                }
            }
            List<Item> imHistoryItems = new List<Item>();
            if (GetAllfolders == false)
            {
                // Get Conversation History items.
                FindItemsResults<Item> findResults;

                ItemView itemView = new ItemView(_pageSize);
                itemView.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                SearchFilter.SearchFilterCollection searchFilterCollection = null;

                do
                {
                    findResults = es.FindItems(_imHistoryFolder.Id, searchFilterCollection, itemView);
                    imHistoryItems.AddRange(findResults);
                    itemView.Offset += _pageSize;
                } while (findResults.MoreAvailable);
            }
            return imHistoryItems;
        }
        
        public Folder FindImHistoryFolder(ExchangeService es, int _pageSize, string FolderName, bool GetAllfolders=false)
        {
            FolderView folderView = new FolderView(_pageSize, 0);
            folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            folderView.PropertySet.Add(FolderSchema.DisplayName);
            folderView.PropertySet.Add(FolderSchema.ChildFolderCount);

            folderView.Traversal = FolderTraversal.Deep;
            Folder imHistoryFolder = null;

            FindFoldersResults findFolderResults;
            bool foundImHistoryFolder = false;
            int count = 0;
            do
            {
                findFolderResults = es.FindFolders(WellKnownFolderName.MsgFolderRoot, folderView);
                foreach (Folder folder in findFolderResults)
                {
                    Program.EmailFolder.Add(folder.DisplayName);
                    if (folder.DisplayName.ToLower() == FolderName && GetAllfolders == false)
                    {
                        imHistoryFolder = folder;
                        foundImHistoryFolder = true;

                    }
                    if (_pageSize==count)
                    {
                        foundImHistoryFolder = true;
                    }
                }
                folderView.Offset += _pageSize;
                count=count+1;
            } while (findFolderResults.MoreAvailable && !foundImHistoryFolder);

            return imHistoryFolder;
        }

        public void QueryImHistory(string[] queryTexts, DateTime createTime, ExchangeService es,int _pageSize)
        {
            char[] participantNames = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            // Get the "Conversation History" folder, if not already found.
            if (_imHistoryFolder == null)
            {
                _imHistoryFolder = this.FindImHistoryFolder(es,_pageSize, "conversation history");
                if (_imHistoryFolder == null)
                {
                    return;
                }
            }

            // Get Conversation History items.
            //_imHistory.Clear();
            FindItemsResults<Item> findResults;

            ItemView itemView = new ItemView(_pageSize);
            itemView.PropertySet = new PropertySet(BasePropertySet.IdOnly);

            // Create a search filter collection, with a logical AND operator, to add query predicates.
            SearchFilter.SearchFilterCollection searchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

            // Add query predicates for conversation topic texts and/or participant names.
            SearchFilter searchFilter = null;

            if (queryTexts != null)
            {
                foreach (string queryText in queryTexts)
                {
                    searchFilter = new SearchFilter.ContainsSubstring(
                        ItemSchema.Body, queryText, ContainmentMode.Substring, ComparisonMode.IgnoreCase);
                    searchFilterCollection.Add(searchFilter);
                }
            }

            if (participantNames != null)
            {
                itemView.PropertySet.Add(ItemSchema.DisplayTo);
                foreach (char pName in participantNames)
                {
                    searchFilter = new SearchFilter.ContainsSubstring(ItemSchema.DisplayTo, pName.ToString());
                    searchFilterCollection.Add(searchFilter);
                }
            }


            // Add the query predicate for the start time of conversations.
            if (createTime != null)
            {
                itemView.PropertySet.Add(ItemSchema.DateTimeCreated);
                searchFilter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeCreated, createTime);
                searchFilterCollection.Add(searchFilter);
            }

            do
            {
                findResults = es.FindItems(_imHistoryFolder.Id, searchFilterCollection, itemView);
                foreach (Item item in findResults.Items)
                {
                    EmailMessage msg = item as EmailMessage;
                    msg.Load(new PropertySet(BasePropertySet.FirstClassProperties));
                    EmailSkypeConvo.Add(msg);
                }
                itemView.Offset += _pageSize;
            } while (findResults.MoreAvailable);

        }
    }
}
