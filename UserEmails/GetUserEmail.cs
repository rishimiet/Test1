using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;
using WinSCP;
using System.Diagnostics;
using System.Xml.Linq;
using System.IO.Compression;

namespace UserEmails
{
    public class GetUserEmail
    {
        public readonly Dictionary<int, Message> messages = new Dictionary<int, Message>();
        public readonly Pop3Client pop3Client;
        //private TreeView listMessages;
        public bool isError = false;

        public GetUserEmail()
        {
            pop3Client = new Pop3Client();
        }

        public void ReceiveMails()
        {
            // Disable buttons while working
            try
            {
                string ftpFileName = "";
                string PII = "",GUID="";
                ArrayList al = new ArrayList();
                bool moveToRP = false;

                //al.Add("wiley.caps@luminad.com:eb9Ap43s");

                //al.Add("cupproduction@luminad.com:y3H6FH68");
                //al.Add("rishiraj.sharma@luminad.com:59D6UaN2");                
                //al.Add("cup.journals@luminad.com:2SLmHG74");
                //al.Add("cup.caps@luminad.com:Tqu2AY4K");

                //al.Add("cupproduction@luminad.com:y3H6FH68");

                //al.Add("rishiraj.sharma@luminad.com:59D6UaN2");                

                // //---------------------------------------------------------------
                // // Live publisher accounts
                // //---------------------------------------------------------------

                al.Add("cup.caps@luminad.com:Tqu2AY4K");
                al.Add("wiley.caps@luminad.com:eb9Ap43s");
                al.Add("cupproduction@luminad.com:y3H6FH68");
                al.Add("aoac.caps@luminad.com:p4jEtRVv");
                al.Add("osa.caps@luminad.com:LWYTdYhZ");
                al.Add("asce.caps@luminad.com:FuGaT583");
                al.Add("hk.caps@luminad.com:x9fW6xUH");
                al.Add("seg.caps@luminad.com:Ya34wdqk");
                al.Add("spie.caps@luminad.com:vN8wpnDP");
                al.Add("ssa.caps@luminad.com:ZWN4cwP7");
                al.Add("aiaa.caps@luminad.com:HH78297m");
                al.Add("csp.caps@luminad.com:ahKMe8Rm");
                al.Add("eeri.caps@luminad.com:xDu29yG8");
                al.Add("astm.caps@luminad.com:kna8mSxK");
                al.Add("cup.journals@luminad.com:2SLmHG74");
                al.Add("NACE.caps@luminad.com:upcCafwr");
                
                //string mailSubjectRead = Console.ReadLine();

                // Live publisher accounts list ends here-------------------------
                //----------------------------------------------------------------

                //al.Add("asme.caps@luminad.com:3Ldkge8T");
                //al.Add("ucp.caps@luminad.com:MvW979kD");
                //al.Add("ha.caps@luminad.com:58HMRLfr");
                //al.Add("aps.caps@luminad.com:3ZuTA34U");                
                int mailCount = 0;
                for (int len = 0; len < al.Count; len++)
                {
                    ftpFileName = "";
                    PII = "";
                    GUID = "";
                    string[] mailDetails = al[len].ToString().Split(':');
                    string mailId = mailDetails[0];
                    string mailPassword = mailDetails[1];
                    
                    if (pop3Client.Connected)
                    {
                        pop3Client.Disconnect();
                    }
                    mailCount = 0;
                   
                    pop3Client.Connect("mail.luminad.com", 995, true);
                                        
                    //pop3Client.Connect("pop.gmail.com", 995, true);
                    //pop3Client.Authenticate("rishiraj.sharma@luminad.com", "59D6UaN2");
                    
                    pop3Client.Authenticate(mailId, mailPassword);

                    Console.WriteLine("Reading Mails from " + mailId + "  ....");
                    int count = pop3Client.GetMessageCount();
                
                    messages.Clear();
                
                    int success = 0;
                    int fail = 0;
                    String sub = "",publisher="";
                    
                    switch (mailId)
                    {
                        case ("aiaa.caps@luminad.com"):
                            publisher = "P1001";
                            break;
                        case ("aoac.caps@luminad.com"):
                            publisher = "P1013";
                            break;
                        case ("asce.caps@luminad.com"):
                            publisher = "P1003";
                            break;
                        case ("astm.caps@luminad.com"):
                            publisher = "P1028";
                            break;
                        case ("csp.caps@luminad.com"):
                            publisher = "P1023";
                            break;
                        case ("eeri.caps@luminad.com"):
                            publisher = "P1010";
                            break;
                        case ("hk.caps@luminad.com"):
                            publisher = "P1027";
                            break;
                        case ("osa.caps@luminad.com"):
                            publisher = "P1005";
                            break;
                        case ("seg.caps@luminad.com"):
                            publisher = "P1006";
                            break;
                        case ("spie.caps@luminad.com"):
                            publisher = "P1007";
                            break;
                        case ("ssa.caps@luminad.com"):
                            publisher = "P1008";
                            break;
                        case ("cup.caps@luminad.com"):
                            publisher = "P1031";
                            break;
                        case ("cup.journals@luminad.com"):
                            publisher = "P1031";
                            break;
                        case ("cupproduction@luminad.com"):
                            publisher = "P1031";
                            break;
                        case ("wiley.caps@luminad.com"):
                            publisher = "P1032";
                            break;
                        case ("NACE.caps@luminad.com"):
                            publisher = "P1030";
                            break;
                        default :
                            publisher = "";
                            break;
                    }

                    
                    DataTable journalDTable = GetDataFromDB("select journalid from ldl_mst_journalInfo  where publisherid='"+publisher+"'");
                    var map = new Dictionary<string, string>();


                    try
                    {
                        if (mailId == "cupproduction@luminad.com" || mailId == "cup.journals@luminad.com" || mailId == "cup.caps@luminad.com")
                        {
                            DataTable dt1 = getSetterMapForANM();
                            foreach (DataRow row in dt1.Rows)
                            {
                                // Setter Number (2) articleid
                                map.Add("ANM-" + row[0].ToString(), row[1].ToString());
                                //Console.WriteLine("ANM-" + row[0].ToString(), row[1].ToString());
                            }
                        }
                    }catch(Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }                   

                    
                    for (int i = count; i >= 1; i -= 1)
                    {
                        try
                        {
                            Message message;
                            try
                            {                                
                                message = pop3Client.GetMessage(i);
                            }
                            catch(Exception c1)
                            {                                
                                Console.WriteLine("Error in reading one of the mail in mailbox: " + mailId + " near " + sub);
                                Console.WriteLine("c1::" + c1.Message);                               
                                continue;
                                //return;
                            }
                            
                            messages.Add(i, message);
                        
                            // Get Sender Name
                            string recFrom = getSenderName(message);     
                            
                        
                            //Working for Time

                            //if (!recFrom.Contains("rishi"))
                            //    continue;


                            string recTime = message.Headers.Date;
                            
                            //var timeUtc = Convert.ToDateTime(message.Headers.Date);
                            //TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");

                            //try
                            //{
                            //    DateTime easternTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, easternZone);
                            //}
                            //catch { }                            

                            bool dayfound = true;
                            try
                            {
                                int num = Convert.ToInt32(recTime.Substring(0, 1));
                                dayfound = false;
                            }
                            catch {}

                            if (dayfound == false)
                            {
                                recTime = "TT " + recTime;
                            }

                            string[] recTimeArr = recTime.Split(' ');
                        
                            //string 
                        
                            // Get Message Subject
                            string msgSubject = getMessageSubject(message);
                            sub = msgSubject;

                           
                            string longsubject = "";
                            bool longsubjectBool = false;
                           
                            // Get mail body
                            System.Net.Mail.MailMessage mailmsg = message.ToMailMessage();                            
                            mailmsg.IsBodyHtml = true;
                            string mb = mailmsg.Body;

                            string mailbody = mailmsg.Body;
                            

                            string fromMailID = mailmsg.From.ToString();
                            string toMailID = mailmsg.To.ToString();
                            string toCC = mailmsg.CC.ToString();
                            MessagePart msgBody = message.FindFirstPlainTextVersion();                           

                            string fromMailAddress = "";
                            try
                            {
                                fromMailAddress = message.Headers.From.Address;
                            }
                            catch { }

                            if (msgBody != null)
                            {
                                mailbody = message.FindFirstPlainTextVersion().GetBodyAsText();
                                try
                                {
                                    mb = message.FindFirstHtmlVersion().GetBodyAsText();
                                    if (mb.Trim() == "")
                                    {
                                        mb = mailbody;
                                    }
                                }
                                catch (Exception e)
                                {
                                    //Console.WriteLine(e.Message);
                                    mb = mailbody;
                                    mb = Regex.Replace(mb, "\\n", "</p><p>");
                                    mb = "<p>" + mb + "</p>";
                                    //Console.WriteLine(e.Message);
                                }
                            }
                            mb = mb.Replace("'", "''");

                            string fileName = recFrom;

                            // testing
                            //if (!mailId.Contains("cup.caps@luminad.com"))
                            //{
                            //if (mailCount > 50)
                            if (mailCount > 350)                                
                            {
                                break;
                            }
                            //}

                            mailCount = mailCount + 1;

                          
                            Console.Write(mailCount + "::");


                            //-- For Testing-------------------------------------------------------------------------
                            //Console.WriteLine("Reading mail:" + mailCount + " ::" + sub);
                            //RDC_FV_S0033822219000614
                            //if (!msgSubject.Contains("AER-1900145 Aero J Paper AeroJ-2019-0077R2"))
                            //{
                            //    continue;
                            //}

                            //if (!mailbody.Contains("PII:"))
                            //{
                            //    continue;
                            //}

                            //publisher = "P1032";
                            //mailId = "wiley.caps@luminad.com";
                            //fromMailID = "cams@cambridge.org";

                            //toMailID = "wiley.caps@luminad.com";
                            ////mailId = "cup.caps@luminad.com";
                            ////fromMailID = "cupproduction.cambridge.org";
                            //fileName = "wileyprod.org";
                            //-------------------------------------Ends here---------------------------------------------------


                            // Get ArticleId from Mail Subject or Mail Body
                            string ArticleID = getArticleID(msgSubject, mailbody, fileName, mailId);

                            string vol = "", issue = "", volIssue = "", issueStageID = "", journID = "", DeliveryType = "", dueDate = "" ;
                            isError = false;
                            moveToRP = false;

                            if (publisher == "P1031" && !fromMailID.ToLower().Contains(".caps") && (mailId.ToLower().Contains("cup.journals@luminad.com") || mailId.ToLower().Contains("cupproduction@luminad.com") || mailId.ToLower().Contains("cup.caps@luminad.com")))
                            {
                                if (sub.Contains("CDRF CAMS Content Delivery Request") && (sub.Contains("FV_") || sub.Contains("EA_")) && sub.StartsWith("LIVE"))
                                {
                                    PII = "";
                                    ArticleID = getArticleID1(msgSubject, mailbody, fileName, mailId);
                                    ArticleID = ArticleID.Replace("PII: ", "");
                                    //Journal Mnemonic: GEO
                                    string jid = getDataByRegex(mailbody, "Journal Mnemonic:\\s[A-Z]+");
                                    jid = jid.Replace("Journal Mnemonic: ", "");
                                    PII = ArticleID.Trim();
                                    ArticleID = PII.Substring(9, 7);
                                    if (jid.Trim() != "")
                                    {
                                        ArticleID = jid.Trim() + "-" + ArticleID;
                                    }
                                    ftpFileName = getDataByRegex(mailbody, "Package: .*?zip");
                                    
                                    ftpFileName = getDataByRegex(ftpFileName, "cams_.*?zip");
                                    //continue;

                                    dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                    dueDate = dueDate.Replace("Due Date : ", "");

                                    if (sub.Contains("ANM") && jid.Trim()=="ANM")
                                    {
                                        string aid = "";
                                        //if (ArticleID.Contains("ANM-"))
                                        //{
                                        //    ArticleID = ArticleID.Replace("ANM-","");
                                        //}
                                        if (map.TryGetValue(ArticleID, out aid))
                                        {
                                            if (aid != "")
                                                ArticleID = "ANM-"+aid;
                                        }
                                        //Console.WriteLine(ArticleID);
                                    }
                                }
                                else if (publisher == "P1031" && (sub.Contains("CDRF CAMS Content Delivery Request") && sub.Contains("_V") && sub.Contains("_I") && sub.StartsWith("LIVE") && sub.Contains(getDataByRegex(sub, "_V.*?_I.*? CDRF CAMS Content Delivery Request"))) && !fromMailID.ToLower().Contains(".caps") && (mailId.ToLower().Contains("cup.journals@luminad.com") || mailId.ToLower().Contains("cupproduction@luminad.com") || mailId.ToLower().Contains("cup.caps@luminad.com")))
                                {
                                    volIssue = getDataByRegex(sub, "_V.*?_I.*? CDRF CAMS Content Delivery Request");
                                    volIssue = getDataByRegex(volIssue, "_V.*?_I.*?\\s");
                                    volIssue = volIssue.TrimStart('_');
                                    string[] arrVI = volIssue.Replace("V","").Replace("I","").Split('_');

                                    string jid = getDataByRegex(mailbody, "Journal Mnemonic:\\s[A-Z]+");
                                    jid = jid.Replace("Journal Mnemonic: ", "");
                                    if (arrVI.Length == 2)
                                    {
                                        vol = arrVI[0];
                                        issue = arrVI[1];
                                    }
                                    volIssue = jid+"_" + volIssue;
                                    issueStageID="ISTM1001";
                                    journID = jid;
                                    DeliveryType = "Issue CAMS Request";
                                    dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                    dueDate = dueDate.Replace("Due Date : ","");

                                    ftpFileName = getDataByRegex(mailbody, "Package: .*?zip");

                                    ftpFileName = getDataByRegex(ftpFileName, "cams_.*?zip");

                                    //Due Date : 03-May-2019 
                                }                               
                                else if (publisher == "P1031" && (sub.Contains("CDRF CAMS Package Import Complete") && sub.Contains("_V") && sub.Contains("_I") && sub.StartsWith("LIVE") && !sub.ToLower().Contains("re-supply") && sub.Contains(getDataByRegex(sub, "_V.*?_I.*? CDRF CAMS Package Import Complete"))) && !fromMailID.ToLower().Contains(".caps") && (mailId.ToLower().Contains("cup.journals@luminad.com") || mailId.ToLower().Contains("cupproduction@luminad.com") || mailId.ToLower().Contains("cup.caps@luminad.com")))
                                {
                                    volIssue = getDataByRegex(sub, "_V.*?_I.*? CDRF CAMS Package Import Complete");
                                    volIssue = getDataByRegex(volIssue, "_V.*?_I.*?\\s");
                                    volIssue = volIssue.TrimStart('_');
                                    string[] arrVI = volIssue.Replace("V", "").Replace("I", "").Split('_');

                                    string jid = getDataByRegex(mailbody, "Journal Mnemonic:\\s[A-Z]+");
                                    jid = jid.Replace("Journal Mnemonic: ", "");
                                    if (arrVI.Length == 2)
                                    {
                                        vol = arrVI[0];
                                        issue = arrVI[1];
                                    }
                                    volIssue = jid + "_" + volIssue;
                                    issueStageID = "ISTM1003";
                                    journID = jid;
                                    DeliveryType = "ISSUE CAMS";
                                    dueDate = "";
                                }
                                else if (publisher == "P1031" && (sub.Contains("CDRF CAMS Package Import Complete") && sub.Contains("_V") && sub.Contains("_I") && sub.StartsWith("LIVE") && sub.Contains("Re-supply CDRF CAMS Package Import Complete") && sub.Contains(getDataByRegex(sub, "_V.*?_I.*?\\s"))) && !fromMailID.ToLower().Contains(".caps") && (mailId.ToLower().Contains("cup.journals@luminad.com") || mailId.ToLower().Contains("cupproduction@luminad.com") || mailId.ToLower().Contains("cup.caps@luminad.com")))
                                {
                                    //volIssue = getDataByRegex(sub, "_V.*?_I.*? CDRF CAMS Package Import Complete");
                                    volIssue = getDataByRegex(sub, "_V.*?_I.*?\\s");
                                    volIssue = volIssue.TrimStart('_');
                                    string[] arrVI = volIssue.Replace("V", "").Replace("I", "").Split('_');

                                    string jid = getDataByRegex(mailbody, "Journal Mnemonic:\\s[A-Z]+");
                                    jid = jid.Replace("Journal Mnemonic: ", "");
                                    if (arrVI.Length == 2)
                                    {
                                        vol = arrVI[0];
                                        issue = arrVI[1];
                                    }
                                    volIssue = jid + "_" + volIssue;
                                    issueStageID = "ISTM1004";
                                    journID = jid;
                                    DeliveryType = "ISSUE CAMS RESUPPLY";
                                    dueDate = "";
                                }                                
                                else if (sub.Contains("CDRF CAMS Package Import Complete"))
                                {
                                    // && !sub.Contains("Re-supply")
                                    dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                    dueDate = dueDate.Replace("Due Date : ", "");

                                    PII = "";
                                    ArticleID = getArticleID1(msgSubject, mailbody, fileName, mailId);
                                    ArticleID = ArticleID.Replace("PII: ", "");
                                    //Journal Mnemonic: GEO
                                    string jid = getDataByRegex(mailbody, "Journal Mnemonic:\\s[A-Z]+");
                                    jid = jid.Replace("Journal Mnemonic: ", "");
                                    PII = ArticleID.Trim();
                                    ArticleID = PII.Substring(9, 7);
                                    if (jid.Trim() != "")
                                    {
                                        ArticleID = jid.Trim() + "-" + ArticleID;
                                    }
                                    ftpFileName = getDataByRegex(mailbody, "Package: .*?zip");
                                    ftpFileName = getDataByRegex(ftpFileName, "cams_.*?zip");
                                    moveToRP = true;

                                    if (sub.Contains("ANM") && jid.Trim() == "ANM")
                                    {
                                        string aid = "";
                                        if (map.TryGetValue(ArticleID, out aid))
                                        {
                                            if (aid != "")
                                                ArticleID = "ANM-" + aid;
                                        }
                                    }
                                }
                                else
                                {
                                    ArticleID = getDataByRegex(msgSubject, "[A-Z]+(\\s)?(-)?[0-9]+");
                                    ArticleID = ArticleID.Replace(" ", "-");
                                    try
                                    {
                                        if (sub.Contains("ANM"))
                                        {
                                            ArticleID = getDataByRegex(msgSubject, "[A-Z]+(\\s)?(-)?[0-9]+((-)[0-9]+)?");
                                            ArticleID = ArticleID.Replace(" ", "-");
                                        }

                                        if (ArticleID.Trim() == "")
                                        {
                                            ArticleID = getDataByRegex(msgSubject.Replace("--", "-"), "[A-Z]+(\\s)?(-)?(\\s)?[0-9]+");
                                            if (ArticleID.Contains("-"))
                                            {
                                                ArticleID = ArticleID.Replace(" ", "");
                                            }
                                            else
                                            {
                                                ArticleID = ArticleID.Replace(" ", "-");
                                            }
                                        }
                                    }
                                    catch { }

                                    if (sub.Contains("ANM"))
                                    {
                                        string aid = "";
                                        //if (ArticleID.Contains("ANM-"))
                                        //{
                                        //    ArticleID = ArticleID.Replace("ANM-", "");
                                        //}
                                        if (map.TryGetValue(ArticleID, out aid))
                                        {
                                            if (aid != "")
                                                ArticleID = "ANM-" + aid;
                                        }
                                    }

                                    if (ArticleID.Trim() == "")
                                    {
                                        //continue;
                                    }
                                    //continue;
                                }
                            }
                            else if (publisher == "P1031" &&  ArticleID.Trim().StartsWith("ANM"))
                            {
                                ArticleID = getDataByRegex(msgSubject, "[A-Z]+(\\s)?(-)?[0-9]+");
                                ArticleID = ArticleID.Replace(" ", "-");
                                try
                                {
                                    if (sub.Contains("ANM"))
                                    {
                                        ArticleID = getDataByRegex(msgSubject, "[A-Z]+(\\s)?(-)?[0-9]+((-)[0-9]+)?");
                                        ArticleID = ArticleID.Replace(" ", "-");
                                    }

                                    if (ArticleID.Trim() == "")
                                    {
                                        ArticleID = getDataByRegex(msgSubject.Replace("--", "-"), "[A-Z]+(\\s)?(-)?(\\s)?[0-9]+");
                                        if (ArticleID.Contains("-"))
                                        {
                                            ArticleID = ArticleID.Replace(" ", "");
                                        }
                                        else
                                        {
                                            ArticleID = ArticleID.Replace(" ", "-");
                                        }
                                    }
                                }
                                catch { }

                                if (sub.Contains("ANM"))
                                {
                                    string aid = "";
                                    if (map.TryGetValue(ArticleID, out aid))
                                    {
                                        if (aid != "")
                                            ArticleID = "ANM-" + aid;
                                    }
                                }
                            }                            

                            if (publisher == "P1031" && (sub.ToUpper().Contains("CDRF CAMS CONTENT DELIVERY IMPORT WAITING") && sub.Contains("_V") && sub.Contains("_I") && sub.StartsWith("LIVE") && sub.Contains(getDataByRegex(sub, "_V.*?_I.*? CDRF CAMS Content Delivery Import Waiting"))) && (mailId.ToLower().Contains("cupproduction@luminad.com")))
                            {
                                //&& !fromMailID.ToLower().Contains(".caps")
                                volIssue = getDataByRegex(sub, "_V.*?_I.*? CDRF CAMS Content Delivery Import Waiting");
                                volIssue = getDataByRegex(volIssue, "_V.*?_I.*?\\s");
                                volIssue = volIssue.TrimStart('_');
                                string[] arrVI = volIssue.Replace("V", "").Replace("I", "").Split('_');

                                string jid = getDataByRegex(sub, "LIVE.*?_");
                                jid = jid.Replace("LIVE ", "").Replace("_", "");
                                if (arrVI.Length == 2)
                                {
                                    vol = arrVI[0];
                                    issue = arrVI[1];
                                }
                                volIssue = jid + "_" + volIssue;
                                issueStageID = "ISTM1002";
                                journID = jid;
                                DeliveryType = "Issue Package Import Waiting";
                                dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                dueDate = dueDate.Replace("Due Date : ", "");
                            }
                            else if (publisher == "P1031" && (sub.ToUpper().Contains("ISSUE APPROVAL") && sub.Contains("_V") && sub.Contains("_I") && sub.StartsWith("LIVE") && sub.Contains(getDataByRegex(sub, "_V.*?_I.*? Issue Approval"))) && ( mailId.ToLower().Contains("cupproduction@luminad.com") ))
                            {
                                //&& !fromMailID.ToLower().Contains(".caps")
                                volIssue = getDataByRegex(sub, "_V.*?_I.*? ISSUE APPROVAL");
                                volIssue = getDataByRegex(volIssue, "_V.*?_I.*?\\s");
                                volIssue = volIssue.TrimStart('_');
                                string[] arrVI = volIssue.Replace("V", "").Replace("I", "").Split('_');

                                string jid = getDataByRegex(sub, "LIVE.*?_");
                                jid = jid.Replace("LIVE ", "").Replace("_", "");
                                if (arrVI.Length == 2)
                                {
                                    vol = arrVI[0];
                                    issue = arrVI[1];
                                }
                                volIssue = jid + "_" + volIssue;
                                issueStageID = "ISTM1005";
                                journID = jid;
                                DeliveryType = "ISSUE APPROVAL";
                                dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                dueDate = dueDate.Replace("Due Date : ", "");
                            }
                            else if (publisher == "P1031" && (sub.ToUpper().Contains("ISSUE SENT FOR APPROVAL") && sub.Contains("_V") && sub.Contains("_I") && sub.StartsWith("LIVE") && sub.Contains(getDataByRegex(sub, "_V.*?_I.*? ISSUE SENT FOR APPROVAL"))) && ( mailId.ToLower().Contains("cupproduction@luminad.com")))
                            {
                                //&& !fromMailID.ToLower().Contains(".caps")
                                volIssue = getDataByRegex(sub, "_V.*?_I.*? ISSUE SENT FOR APPROVAL");
                                volIssue = getDataByRegex(volIssue, "_V.*?_I.*?\\s");
                                volIssue = volIssue.TrimStart('_');
                                string[] arrVI = volIssue.Replace("V", "").Replace("I", "").Split('_');

                                string jid = getDataByRegex(sub, "LIVE.*?_");
                                jid = jid.Replace("LIVE ", "").Replace("_", "");
                                if (arrVI.Length == 2)
                                {
                                    vol = arrVI[0];
                                    issue = arrVI[1];
                                }
                                volIssue = jid + "_" + volIssue;
                                issueStageID = "ISTM1006";
                                journID = jid;
                                DeliveryType = "ISSUE SENT FOR APPROVAL";
                                dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                dueDate = dueDate.Replace("Due Date : ", "");
                            }
                            // End here                           

                            if (publisher == "P1032" && !fromMailID.ToLower().Contains(".caps") && mailId.ToLower().Contains("wiley.caps@luminad.com"))
                            {
                                if (sub.Contains("Task Assignment - RT Copyediting for") && mailbody.Contains("Dear Lumina Typesetter"))
                                {
                                    string jid = getDataByRegex(mailbody, "The production task GUID is: {.*?}");
                                    jid = jid.Replace("The production task GUID is: {", "");
                                    jid = jid.Replace("}", "");
                                    GUID = jid;                                    
                                }
                            }                           
                            // Getting GUID from Wiley Mail
                            if(publisher=="P1032" && !fromMailID.ToLower().Contains(".caps") && mailId.ToLower().Contains("wiley.caps@luminad.com") && sub.Contains("Task Assignment -"))
                            {
                                string jid = getDataByRegex(mailbody, "The production task GUID is: {.*?}");
                                jid = jid.Replace("The production task GUID is: {", "");
                                jid = jid.Replace("}", "");
                                GUID = jid;
                            }

                            string dirStartPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString();

                            // Working for Revision Article
                            if (ArticleID.ToLower().EndsWith("r") && mailId.Contains("asce.caps"))
                            {
                                String[] dirs = Directory.GetDirectories(dirStartPath, "LDL-ASCE-*-" + ArticleID);
                                if (dirs.Length == 0)
                                {
                                    ArticleID = ArticleID.TrimEnd('R');
                                    ArticleID = ArticleID.TrimEnd('r');
                                }
                            }

                            // if message length is greater then 150 then save subject with 50 charcters for filename
                            if (msgSubject.Length > 150)
                            {
                                longsubject = msgSubject;
                                msgSubject = msgSubject.Substring(0, 50);
                                msgSubject = msgSubject + "....";
                                longsubjectBool = true;
                            }
                            
                            //msgSubject = msgSubject.Replace("--", "");
                            // Creating Path

                            //string dirStartPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString();
                            if (!dirStartPath.EndsWith(@"\"))
                            {
                                dirStartPath = dirStartPath + @"\";
                            }

                            string fileToSave = recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + recTimeArr[4].Replace(':', '.') + "_ldps_" + msgSubject;
                            fileToSave = fileToSave.Replace("	", "");

                            string mailRecTime = recTimeArr[4].Replace(':', '.');
                            string[] arrMailRecTime = mailRecTime.Split('.');

                            if (ArticleID == "")
                            {
                                ArticleID = "NotFound";
                            }

                            ArticleID = ArticleID.Replace(".", "-");                            
                            string foldToSearchIn = "";
                            try
                            {
                                string[] foldNameArr = mailId.Split('.');
                                foldToSearchIn = foldNameArr[0];

                                if (mailId.ToLower().Contains("cupproduction"))
                                {
                                    foldToSearchIn = "CUP";
                                }
                            }
                            catch { }

                            // CAMS MAIL RECEIVED FOR ACCEPTED MANUSCRIPT
                            if (sub.Contains("CDRF CAMS Content Delivery Request") && (sub.Contains("_AM_")) && sub.StartsWith("LIVE"))
                            {
                                
                                ArticleID = getArticleIDForAM(msgSubject, mailbody, fileName, mailId, sub,map);
                            }

                            // Get the directory location for particular article.
                            ArticleID = GetFolderCode(dirStartPath,foldToSearchIn,ArticleID);

                            string wileyArticleID = ArticleID;
                            if (!ArticleID.Trim().StartsWith("LDL-"))
                            {
                                ArticleID = "NotFound";
                            }
                            
                            // Set outbox folder location if received from caps
                            string mailFolder = "";
                            string notfoundFolder = "";
                            if (fileName.Contains("spie.caps") || fileName.Contains("aoac.caps") || fileName.Contains("asme.caps") || fileName.Contains("ssa.caps")
                                || fileName.Contains("aiaa.caps") || fileName.Contains("osa.caps") || fileName.Contains("asce.caps") || fileName.Contains("NACE.caps") || fileName.Contains("nace.caps")
                                //|| fileName.Contains("wiley.caps")
                                || fileName.Contains("ucp.caps") || fileName.Contains("ha.caps") || fileName.Contains("seg.caps") || fileName.Contains("cup.caps")
                                || fileName.Contains("aps.caps") || fileName.Contains("hk.caps") || fileName.Contains("astm.caps") || fileName.Contains("eeri.caps") || fileName.Contains("aoac.caps") || fileName.Contains("wiley.caps"))
                            {
                                //Console.WriteLine("outbox mail articleID: " + ArticleID);
                                mailFolder = @"\mail\OUTBOX";
                                notfoundFolder = @"\mail\OUTBOX\" +foldToSearchIn + @"\";
                            }
                            else
                            {
                                try
                                {                                    
                                    string query = "SELECT at.publisherid,at.journalid,at.articleid FROM ldl_trn_articletracker at left join ldl_mst_publisher p on p.publisherid=at.publisherid  where FOLDERCODE='" + ArticleID + "'";
                                    DataTable dt1 = GetDataFromDB(query);
                                    if (dt1.Rows.Count == 1)
                                    {
                                        DataRow dr1 = dt1.Rows[0];
                                        string publisherID = dr1[0].ToString();
                                        string journalID = dr1[1].ToString();
                                        string articleID = dr1[2].ToString();

                                        dt1 = GetDataFromDB("Select * from ldl_trn_mailDetails where publisherid='" + publisherID + "' and journalid='" + journalID + "' and articleid='" + articleID + "' and MSG_RECEIVE_TIME='" + recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + " " + recTimeArr[4] + "'");
                                        if (dt1.Rows.Count == 0)
                                        {
                                            string dbQuery = "insert into ldl_trn_mailDetails(publisherid,journalid,articleid,MSG_RECEIVE_TIME,mailstatus) values('" + publisherID + "' ,'" + journalID
                                                + "' ,'" + articleID + "', '" + recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + " " + recTimeArr[4] + "','1')";
                                            UpdateQuery(dbQuery);
                                        }
                                    }
                                }
                                catch { }
                                mailFolder = @"\mail";
                                notfoundFolder = @"\mail\" + foldToSearchIn + @"\";
                            }


                            //6-05-2019
                            //string dirPath = dirStartPath + ArticleID + @"\mail" + @"\" + fileName + @"\";
                            string dirPath = dirStartPath + ArticleID + mailFolder + @"\" + fileName + @"\";
                            if (volIssue != "" & vol.Trim()!= "" && issue != "")
                            {
                                if(!Directory.Exists(dirStartPath +"CUPIssueMail\\" + volIssue)){
                                    Directory.CreateDirectory(dirStartPath + "CUPIssueMail\\" + volIssue);
                                    }
                                dirPath = dirStartPath + "CUPIssueMail\\" + volIssue + "\\" + mailFolder + @"\" + fileName + @"\";
                                dirStartPath = dirStartPath + "CUPIssueMail\\";
                                ArticleID = volIssue;
                            }
                            
                            if (!Directory.Exists(dirStartPath + ArticleID) || ArticleID == "NotFound")
                            {                                
                                string journalName = "";
                                for (int jd = 0; jd < journalDTable.Rows.Count; jd++)
                                {
                                    DataRow jdT1 = journalDTable.Rows[jd];
                                    string journal = jdT1[0].ToString();
                                    if (mailbody.Contains(" " + journal + " ") || mailbody.Contains("_" + journal + "_") || mailbody.Contains(" " + journal + "-") || mailbody.Contains("-" + journal + "-") || mailbody.Contains("-" + journal + " ") || mailbody.Contains(journal + " "))
                                    {
                                        journalName = journal + "\\";
                                        break;
                                    }
                                    if (msgSubject.Contains(" " + journal + " ") || msgSubject.Contains("_" + journal + "_") || msgSubject.Contains(" " + journal + "-") || msgSubject.Contains("-" + journal + "-") || msgSubject.Contains("-" + journal + " ") ||  msgSubject.Contains(journal + " "))
                                    {
                                        journalName = journal + "\\";
                                        break;
                                    }

                                    if ((msgSubject.Contains(journal + "ENG") || mailbody.Contains(journal + "ENG")) && publisher == "P1003")
                                    {
                                        journalName = journal + "\\";
                                        break;
                                    }
                                }

                                if (journalName == "")
                                    journalName = "UnIdentified\\";

                                if (fileName.ToLower().Contains(".caps"))
                                {
                                    dirPath = dirStartPath + "NotFound" + "\\" + foldToSearchIn.ToUpper() + "\\" + journalName + "OUTBOX\\" + fileName + @"\";
                                }
                                else
                                {
                                    dirPath = dirStartPath + "NotFound" + "\\" + foldToSearchIn.ToUpper() + "\\" + journalName + "INBOX\\"  + fileName + @"\";
                                }
                            }

                            if (!Directory.Exists(dirPath))
                            {                                
                                Directory.CreateDirectory(dirPath);
                            }

                            if (ArticleID == "")
                            {
                                ArticleID = "NotFound";
                            }                         

                            // Working for Attachement
                            string attachmentToForward = "";

                            string attTime = recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + recTimeArr[4].Replace(':', '.')+ "_ldps_";
                            foreach (MessagePart mpart in message.FindAllAttachments())
                            {
                                string fileNameAttach = mpart.FileName;
                                string file_name_attach = mpart.ContentType.MediaType;                                

                                try
                                {
                                    FileInfo finfo;
                                    if (fileNameAttach.Replace("#", "").Length > 70)
                                    {
                                        int lent = fileNameAttach.Length;
                                        finfo = new FileInfo(dirPath + attTime + fileNameAttach.Replace("#", "").Substring(0, 57)+"_" + fileNameAttach.Replace("#", "").Substring(lent - 4));
                                    }
                                    else
                                    {
                                        finfo = new FileInfo(dirPath + attTime + fileNameAttach.Replace("#", ""));
                                    }                                    


                                    mpart.Save(finfo);

                                    if (!fileNameAttach.ToLower().Contains("winmail"))
                                    {
                                        if (File.Exists(finfo.FullName))
                                        {
                                            attachmentToForward = attachmentToForward + ";" + finfo.FullName;
                                        }
                                    }
                                }
                                catch(Exception e) {
                                    Console.WriteLine(e.Message);
                                }
                               

                                if (publisher == "P1031" && !sub.StartsWith("RE:") && ArticleID != "NotFound")
                                {
                                    //CUP PO PDF and XML 
                                    // Samir has confirmed that tysetting PO is confirmed receving 8-aug-2019

                                    if (fileNameAttach.ToLower().Contains("po_") && !fileNameAttach.ToLower().Contains("pre") && fileNameAttach.ToLower().Contains(".pdf"))
                                    {
                                        if (!Directory.Exists(dirStartPath + ArticleID + "\\CompareMeta") && publisher == "P1031")
                                        {
                                            Directory.CreateDirectory(dirStartPath + ArticleID + "\\CompareMeta");
                                           
                                        }
                                        if (!File.Exists(dirStartPath + ArticleID + "\\CompareMeta\\" + fileNameAttach))
                                        {

                                            FileInfo finfoo = new FileInfo(dirStartPath + ArticleID + "\\CompareMeta\\" + fileNameAttach.Replace("#", ""));
                                            mpart.Save(finfoo);
                                            
                                            UpdateInArticleMovementRegister(ArticleID,getCurrentStageAndName(publisher, ArticleID),"Mail received for PO PDF and XML");

                                            JMSProducer jProd = new JMSProducer();
                                            jProd.postMessageforCompareMeta(ArticleID, "COMPARE");
                                        }
                                    }
                                }
                            }

                            attachmentToForward = attachmentToForward.TrimStart(';');
                            string path = dirPath + fileToSave + ".mail";
                            string emlPath = dirPath + fileToSave + ".eml";
                            // Saving Mail Body

                            //Console.WriteLine("Mail Path:: " + path);                         

                            
                            if (!File.Exists(path))
                            {                                
                                // If File already Exist with same name in one hour than do copy in CAPSMail Folder.
                                try
                                {
                                    if (recFrom.ToLower().EndsWith(".caps"))
                                    {
                                        string[] files = System.IO.Directory.GetFiles(dirPath, recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + arrMailRecTime[0] + "." + "*_ldps_" + msgSubject + ".mail", System.IO.SearchOption.AllDirectories);

                                        bool MailTofound;
                                        MailTofound = false;

                                        if (files.Length > 0)
                                        {
                                            string txt = "To  : " + toMailID;
                                            for (int f = 0; f < files.Length; f++)
                                            {
                                                string fileText = File.ReadAllText(files[f]);
                                                if (fileText.Contains(txt))
                                                {
                                                    MailTofound = true;
                                                    break;
                                                }
                                            }
                                        }

                                        if (MailTofound == true)
                                        {
                                            if (files.Length > 0)
                                            {
                                                path = "D:\\CAPSMail\\" + fileToSave + ".mail";
                                            }
                                            else
                                            {
                                                files = System.IO.Directory.GetFiles(dirPath, "0" + recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + arrMailRecTime[0] + "." + "*_ldps_" + msgSubject + ".mail", System.IO.SearchOption.AllDirectories);
                                                if (files.Length > 0)
                                                {
                                                    path = "D:\\CAPSMail\\" + fileToSave + ".mail";
                                                }
                                            }
                                        }
                                    }
                                }
                                catch { }

                                if (sub.Contains("CDRF CAMS Content Delivery Request") && !fromMailID.ToLower().Contains("lumina") && (sub.Contains("_AM_")) && sub.StartsWith("LIVE") && ArticleID.Trim().StartsWith("LDL-"))
                                {                                    
                                    dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                    dueDate = dueDate.Replace("Due Date : ", "");
                                    try
                                    {
                                        string ftpAMFileName = getDataByRegex(mailbody, "Package: .*?zip");
                                        ftpAMFileName = getDataByRegex(ftpAMFileName, "cams_.*?zip");
                                        string dPathAM = dirStartPath + ArticleID + "\\";
                                        if (!Directory.Exists(dPathAM + "\\AM\\FROM CUP\\"))
                                        {
                                            Directory.CreateDirectory(dPathAM + "\\AM\\FROM CUP\\");
                                        }
                                        try
                                        {
                                            string[] filesAM = Directory.GetFiles(dPathAM + "\\AM\\FROM CUP\\", "*.zip", SearchOption.TopDirectoryOnly);
                                            if (filesAM.Length > 0)
                                            {
                                                if (!Directory.Exists(dPathAM + "\\AM_Backup"))
                                                {
                                                    Directory.CreateDirectory(dPathAM + "\\AM_Backup");
                                                }
                                                foreach (String fname in filesAM)
                                                {
                                                    FileInfo f = new FileInfo(fname);
                                                    f.CopyTo(dPathAM + "\\AM_Backup\\" + f.Name, true);
                                                }
                                            }
                                        }
                                        catch { }
                                        
                                        downloadZIPFromFTP("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", dPathAM + "\\AM\\FROM CUP\\" + ftpAMFileName, "//from_cams//" + ftpAMFileName, "none");
                                        updateForAM(ArticleID, sub, dueDate);
                                        JMSProducer jProd = new JMSProducer();
                                        jProd.postAMPDFMessage(ArticleID, "AM PDF Request");
                                        try
                                        {
                                            Process processChild = Process.Start(@"D:\executor_portal\CupCamsProcess\CupCamsProcess.bat", ArticleID);
                                            processChild.WaitForExit(10000);
                                            processChild.Kill();
                                        }
                                        catch(Exception e) {
                                            Console.WriteLine(e.Message);
                                        }

                                        string outputFileName;
                                        string query = "SELECT ad.PII_New4Display FROM ldl_trn_articletracker at left join ldl_trn_articledetails ad on at.journalid=ad.journalid and at.articleid=ad.articleid where FOLDERCODE='" + fileName + "'";
                                        string am_PII = "", doi = "";
                                        System.Data.DataTable dt1 = GetDataFromDB(query);
                                        if (dt1.Rows.Count != 0)
                                        {
                                            DataRow dr1 = dt1.Rows[0];
                                            doi = dr1[0].ToString();
                                            am_PII = dr1[0].ToString() + 'a';
                                        }
                                        else
                                        {
                                            am_PII = "a";
                                        }
                                        var foldersFound = Directory.GetDirectories(dPathAM + "\\AM\\TO CUP\\", "WebPDF", SearchOption.AllDirectories);
                                        if (foldersFound.Length == 0)
                                        {
                                            outputFileName = dPathAM + "\\AM\\" + am_PII + ".pdf";
                                        }
                                        else
                                        {
                                            outputFileName = foldersFound[0] + "\\" + am_PII + ".pdf";
                                        }
                                        if(File.Exists(dPathAM + "\\AM\\" + am_PII + ".pdf") && foldersFound.Length>0){
                                            File.Copy(dPathAM + "\\AM\\" + am_PII + ".pdf", outputFileName);
                                        }
                                    }
                                    catch { }
                                    
                                }
                                else if (sub.Contains("CDRF CAMS Content Delivery Request") && !fromMailID.ToLower().Contains("lumina") && (sub.Contains("_AM_")) && sub.StartsWith("LIVE") && !ArticleID.Trim().StartsWith("LDL-"))
                                {
                                    try {
                                        string ftpAMFileName = getDataByRegex(mailbody, "Package: .*?zip");
                                        ftpAMFileName = getDataByRegex(ftpAMFileName, "cams_.*?zip");
                                        downloadZIPFromFTP("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", "D:\\CAPS\\CUP_AM\\" + ftpAMFileName, "//from_cams//" + ftpAMFileName, "none");

                                    }
                                    catch { }
                                    
                                }

                                if (!File.Exists(path))
                                {
                                    //if (publisher == "P1032" && !recFrom.ToLower().EndsWith(".caps"))
                                    //{
                                    //    string mailDetail = "";
                                    //    try
                                    //    {
                                    //        if (longsubjectBool == true)
                                    //        {
                                    //            mailDetail = "<p>Subject: " + longsubject + "</p><br/>";
                                    //        }
                                    //        if (!fromMailID.Contains("@") && fromMailAddress.Trim().Contains("@"))
                                    //        {
                                    //            mailDetail = mailDetail + "<p>" + "From: " + fromMailAddress + "</p>";
                                    //        }
                                    //        else
                                    //        {
                                    //            mailDetail = mailDetail + "<p>" + "From: " + fromMailID + "</p>";
                                    //        }
                                    //        mailDetail = mailDetail + "<p>" + "To  : " + toMailID + "</p>";
                                    //        if (toCC != "" && toCC != fromMailID)
                                    //        {
                                    //            mailDetail = mailDetail + "<p>" + "CC  : " + toCC + "</p>";
                                    //        }
                                    //        mailDetail = mailDetail + "<p>Sent: " + recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + " " + recTimeArr[4].Replace(':', '.');
                                    //        mailDetail = mailDetail.Replace("'", "''");
                                    //    }
                                    //    catch { }

                                    //    string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + mailDetail + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate())";
                                    //    UpdateQuery(dbQuery);
                                    //}


                                    // GET Current Stage
                                    string currentStage = getCurrentStageAndName(publisher, ArticleID);

                                    StreamWriter sw ;
                                    try
                                    {                                       
                                        if (publisher == "P1032" && !recFrom.ToLower().EndsWith(".caps"))
                                        {
                                            string mailDetail = "";
                                            try
                                            {
                                                if (longsubjectBool == true)
                                                {
                                                    mailDetail = "<p>Subject: " + longsubject + "</p><br/>";
                                                }
                                                if (!fromMailID.Contains("@") && fromMailAddress.Trim().Contains("@"))
                                                {
                                                    mailDetail = mailDetail + "<p>" + "From: " + fromMailAddress + "</p>";
                                                }
                                                else
                                                {
                                                    mailDetail = mailDetail + "<p>" + "From: " + fromMailID + "</p>";
                                                }
                                                mailDetail = mailDetail + "<p>" + "To  : " + toMailID + "</p>";
                                                if (toCC != "" && toCC != fromMailID)
                                                {
                                                    mailDetail = mailDetail + "<p>" + "CC  : " + toCC + "</p>";
                                                }
                                                mailDetail = mailDetail + "<p>Sent: " + recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + " " + recTimeArr[4].Replace(':', '.');
                                                mailDetail = mailDetail.Replace("'", "''");
                                            }
                                            catch { }

                                            try
                                            {
                                                DataTable dt1 = GetDataFromDB("Select * from caps_trn_MailTracker where tomailid='vch.production@luminad.com' and mailsubject='" + msgSubject.Replace("'", "''") + "'");
                                                if (dt1.Rows.Count > 0)
                                                {
                                                }
                                                else
                                                {
                                                    string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + mailDetail + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate())";
                                                    UpdateQuery(dbQuery);
                                                }
                                            }
                                            catch
                                            {
                                                string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + mailDetail + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate())";
                                                UpdateQuery(dbQuery);
                                            }                                            
                                        }

                                        // Start here for wiley
                                        bool isWileyArticleAlreadyIntegerated = false;
                                        if (GUID != "" && publisher == "P1032")
                                        {
                                            isWileyArticleAlreadyIntegerated=generateWileyMailorDownload(mailbody, mailId, wileyArticleID, publisher, fromMailID, sub, GUID, msgSubject, mb, dirStartPath);
                                            if (isWileyArticleAlreadyIntegerated == false)
                                            {
                                                continue;
                                            }
                                        }

                                        if (sub.Contains("Task Assignment") && ArticleID.ToUpper().Contains("LDL-WILEY-") && mailbody.Contains("Typesetter") && mailbody.ToLower().Contains("rt accepted article") && mailId.ToLower().Contains("wiley.caps@luminad.com") && !fromMailID.ToLower().Contains("lumina"))
                                        {
                                            string downloadSuccess = downloadWileyAMPackage(mailbody, mailId, wileyArticleID, publisher, fromMailID, sub, GUID, msgSubject, mb, dirStartPath);
                                            if (downloadSuccess != "success")
                                            {
                                                continue;
                                            }
                                        }
                                        else if (sub.Contains("Task Assignment") && publisher == "P1032" && mailbody.Contains("Typesetter") && mailbody.Contains("RT Accepted Article") && mailId.ToLower().Contains("wiley.caps@luminad.com") && !fromMailID.ToLower().Contains("lumina"))
                                        {
                                            if (!ArticleID.ToUpper().Contains("LDL-WILEY"))
                                            {

                                                 string[] copyEditFilesCount = Directory.GetFiles(dirStartPath + "\\NotFound\\WILEY\\UnIdentified\\INBOX\\em", "*" + ArticleID.Replace("-",".") + "*.mail");

                                                 if (copyEditFilesCount.Length > 0)
                                                 {
                                                     continue;
                                                 }

                                                string jid = getDataByRegex(mailbody, "The production task GUID is: {.*?}");
                                                jid = jid.Replace("The production task GUID is: {", "");
                                                jid = jid.Replace("}", "").Trim();
                                                GUID = jid.Trim();

                                                string downloadSuccess = downloadWileyAMPackage(mailbody, mailId, wileyArticleID, publisher, fromMailID, sub, GUID, msgSubject, mb, dirStartPath);

                                                if (downloadSuccess.ToLower().Contains("fail"))
                                                {
                                                    UpdateQuery("INSERT INTO caps_trn_MailTracker (PublisherID, FromMailID, ToMailID, MailSubject, MailBody) VALUES ('P1032', 'wiley.caps@luminad.com', 'rishiraj.sharma@luminad.com;sandeep.kumar@luminad.com', '" + downloadSuccess + "', '<p>Hi Team,</p><p>Please check for " + sub.Replace("'", "") + " integeration fail.</p><p></p><p>Regards,</p><p>CAPSTeam</p>');");
                                                    continue;
                                                }
                                                //continue;
                                            }
                                        }
                                        // end here for wiley

                                        // For CUP Issue
                                        if (volIssue != "" & vol.Trim() != "" && issue != "")
                                        {
                                            try
                                            {
                                                Console.WriteLine("Start for CAMS:" + volIssue);

                                                string submittedStage = "";
                                                if (issueStageID == "ISTM1007")
                                                {
                                                    submittedStage = "ISTM1006";
                                                }
                                                else if (issueStageID == "ISTM1002")
                                                {
                                                    submittedStage = "ISTM1001";
                                                }
                                                else if (issueStageID == "ISTM1004")
                                                {
                                                    submittedStage = "ISTM1004";
                                                }

                                                if (submittedStage != "")
                                                {
                                                    UpdateQuery(" Update ldl_trn_CUP_IssueCAMSDetails set IssueStageOutTime=getdate() where  VolumeIssueNumber='" + volIssue + "' and journalid='" + journID + "' and IssueStageOutTime is null and IssueStageID='" + submittedStage + "'");
                                                }                                                

                                                if (dueDate != "")
                                                {
                                                    UpdateQuery("insert into ldl_trn_CUP_IssueCAMSDetails (PublisherID,JournalID,Volume,Issue,VolumeIssueNumber,MailTime,DeliveryType,IssueDueDate,IssueStageID) values ('P1032','" + journID + "','" + vol + "','" + issue + "','" + volIssue + "',getdate(),'" + DeliveryType + "','" + dueDate + "','" + issueStageID + "')");
                                                }
                                                else
                                                {
                                                    UpdateQuery("insert into ldl_trn_CUP_IssueCAMSDetails (PublisherID,JournalID,Volume,Issue,VolumeIssueNumber,MailTime,DeliveryType,IssueStageID) values ('P1032','" + journID + "','" + vol + "','" + issue + "','" + volIssue + "',getdate(),'" + DeliveryType + "','" + issueStageID + "')");
                                                }

                                                if (DeliveryType == "Issue CAMS Request" && dueDate != "")
                                                {
                                                    UpdateQuery(" UPDATE ldl_trn_IssueVolume SET ISSUESTAGEID='" + issueStageID + "',IssueDueDate='" + dueDate + "' WHERE PublisherID='P1031' AND JournalID='" + journID + "' AND VolNo='" + vol + "' AND IssueNo='" + issue + "';");
                                                }
                                                else
                                                {
                                                    UpdateQuery(" UPDATE ldl_trn_IssueVolume SET ISSUESTAGEID='" + issueStageID + "' WHERE PublisherID='P1031' AND JournalID='" + journID + "' AND VolNo='" + vol + "' AND IssueNo='" + issue + "';");
                                                }

                                                if (issueStageID.Trim() == "ISTM1001")
                                                {
                                                    string dPath = dirStartPath + ArticleID.Trim() + "\\";
                                                    if (!Directory.Exists(dPath + "\\ISSUECAMSDelivery\\FROM CUP"))
                                                    {

                                                        Directory.CreateDirectory(dPath + "\\ISSUECAMSDelivery\\");
                                                    }
                                                    downloadZip("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", dPath + "\\ISSUECAMSDelivery\\FROM CUP\\" + ftpFileName, "//from_cams//" + ftpFileName, "none", "");

                                                    if (File.Exists(dPath + "\\ISSUECAMSDelivery\\FROM CUP\\" + ftpFileName))
                                                    {
                                                        JMSProducer jProd = new JMSProducer();
                                                        jProd.postMessageforIssueCAMS("P1031:" + journID + ":" + vol + ":" + issue + ":issueCamsBundle", "");
                                                    }

                                                }

                                            }
                                            catch { }
                                        }
                                        if (publisher == "P1031" && (fromMailID.ToLower().Contains("anmproduction")) && (sub.Contains("ANM PO & XML") || sub.Contains("ANM PO & XML") || sub.Contains("ANM PO")))
                                        {
                                            Console.WriteLine("Received ANM Mail");
                                            bool havePDF = false;
                                            bool haveXML = false;
                                            
                                            foreach (MessagePart mpart in message.FindAllAttachments())
                                            {
                                                string fileNameAttach = mpart.FileName;
                                                if (fileNameAttach.ToLower().Contains("anm_po_for_copy_editing_setting.pdf"))
                                                {
                                                    havePDF = true;
                                                }
                                                if (fileNameAttach.ToLower().Contains("xml_export_anm_") && fileNameAttach.ToLower().Contains(".xml"))
                                                {
                                                    haveXML = true;
                                                }
                                            }

                                            if (haveXML == true && havePDF == true)
                                            {
                                                Console.WriteLine("Received ANM Mail copying started.");
                                                try
                                                {
                                                    String sDate = DateTime.Now.ToString();
                                                    DateTime datevalue = (Convert.ToDateTime(sDate.ToString()));

                                                    String dy = datevalue.Day.ToString();
                                                    String mn = datevalue.Month.ToString();
                                                    String yy = datevalue.Year.ToString();

                                                    string dir = "D:\\CAPS\\ANM-INTEGERATION\\INCOMING" + "\\" + dy + "-" + mn + "-" + yy;
                                                    string[] dirCount = Directory.GetDirectories("D:\\CAPS\\ANM-INTEGERATION\\INCOMING", dy + "-" + mn + "-" + yy + "*");
                                                    dir = dir + "_" + (dirCount.Length + 1).ToString();
                                                    if (!Directory.Exists(dir))
                                                    {
                                                        Directory.CreateDirectory(dir);
                                                    }
                                                    foreach (MessagePart mpart in message.FindAllAttachments())
                                                    {
                                                        string fileNameAttach1 = mpart.FileName;
                                                        if (fileNameAttach1.ToLower().Contains("anm_po_for_copy_editing_setting.pdf"))
                                                        {
                                                            FileInfo finfo = new FileInfo(dir + "\\" + fileNameAttach1.Replace("#", ""));
                                                            mpart.Save(finfo);
                                                        }
                                                        if (fileNameAttach1.ToLower().Contains("xml_export_anm"))
                                                        {
                                                            FileInfo finfo = new FileInfo(dir + "\\" + fileNameAttach1.Replace("#", ""));
                                                            mpart.Save(finfo);
                                                        }
                                                    }
                                                    try
                                                    {
                                                        JMSProducer jProd = new JMSProducer();
                                                        jProd.postMessage(new System.IO.DirectoryInfo(dir).Name, "ANM Article Integeration");
                                                    }
                                                    catch(Exception e) {
                                                        Console.WriteLine(e.Message);
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                       
                                        //sw = new StreamWriter(path, true, Encoding.UTF8);
                                        if (!path.Contains("NotFound\\") && !sub.Contains("_AM_") && sub.Contains("CAMS Content Delivery") && !fromMailID.ToLower().Contains("lumina") && path.Contains("LDL-CUP-") && !recFrom.ToLower().Contains("caps") && PII.Trim() != "" && (mailId.ToLower().Contains("cup.journals@luminad.com") || mailId.ToLower().Contains("cupproduction@luminad.com")))
                                        {
                                            
                                            if (!path.ToLower().Contains("notfound") && moveToRP==false)
                                            {
                                                string dPath = dirStartPath + ArticleID +"\\";
                                                if (!Directory.Exists(dPath + "\\CAMS Delivery\\"))
                                                {

                                                    Directory.CreateDirectory(dPath + "\\CAMS Delivery\\");
                                                }
                                                else
                                                {
                                                    if (sub.ToLower().Contains("re-supply"))
                                                    {
                                                        try
                                                        {
                                                            Directory.Move(dPath + "\\CAMS Delivery\\", dPath + "\\CAMS Delivery" + recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + recTimeArr[4].Replace(':', '.'));
                                                            if (!Directory.Exists(dPath + "\\CAMS Delivery\\"))
                                                            {
                                                                Directory.CreateDirectory(dPath + "\\CAMS Delivery\\");
                                                            }
                                                        }
                                                        catch { }
                                                    }                                                    
                                                }

                                                if (!Directory.Exists(dPath + "\\CAMS Delivery\\FROM CUP\\"))
                                                {                                                   
                                                    Directory.CreateDirectory(dPath + "\\CAMS Delivery\\FROM CUP\\");                                                    
                                                    Console.WriteLine("Create directory");

                                                }
                                                
                                                if (!Directory.Exists(dPath + "\\CAMS Delivery\\TO CUP\\"))
                                                {
                                                    Directory.CreateDirectory(dPath + "\\CAMS Delivery\\TO CUP\\");
                                                }

                                                DataTable dt1 = GetDataFromDB("select t.stageid,t.journalid,t.articleid,t.INPUTTYPE from ldl_trn_articletracker t  where  foldercode='" + ArticleID + "'");
                                                if (!File.Exists(dPath + "\\CAMS Delivery\\FROM CUP\\" + ftpFileName) && dPath.Contains(ArticleID))
                                                {                                                    
                                                    try
                                                    {
                                                        if (!sub.ToLower().Contains("re-supply"))
                                                        {
                                                            if (dt1.Rows.Count == 1)
                                                            {
                                                                DataRow dr1 = dt1.Rows[0];
                                                                string stageID = dr1[0].ToString();
                                                                if (stageID == "STM1012" || stageID == "STM1032")
                                                                {
                                                                    createLog(dPath, "Start downloading package for CAMS request received.");
                                                                    downloadZip("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", dPath + "\\CAMS Delivery\\FROM CUP\\" + ftpFileName, "//from_cams//" + ftpFileName, "none", dPath);
                                                                    Console.WriteLine("CAMS:" + sub);                                                                   
                                                                }
                                                                else
                                                                {
                                                                    createLog(dPath, "Package now downloaded as article is not in Author and nor in PE");
                                                                    try
                                                                    {
                                                                        if (!Directory.Exists(dPath + "\\CAMSDeliveryForOtherStage\\FROM CUP\\"))
                                                                        {
                                                                            Directory.CreateDirectory(dPath + "\\CAMSDeliveryForOtherStage\\FROM CUP\\");
                                                                        }
                                                                        downloadZip("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", dPath + "\\CAMSDeliveryForOtherStage\\FROM CUP\\" + ftpFileName, "//from_cams//" + ftpFileName, "none", dPath);
                                                                        bool result = moveProcessedFile("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", "/from_cams/Processed//" + ftpFileName, "/from_cams/" + ftpFileName, "none");

                                                                        string attach = dPath + "\\CAMSDeliveryForOtherStage\\FROM CUP\\" + ftpFileName;

                                                                        UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','kolapparaja.k@luminad.com;cup.pm@luminad.com;sandeep.sharda@luminad.com;manish.bansal@luminad.com;cup.xmlpdy@luminad.com','capssupport@luminad.com','CAMS received for article not in PE stage and nor in Author: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team as no action taken for CAMS request received.</p><br/>" + mb + "',getdate(),'" + attach + "')");

                                                                    }
                                                                    catch { }
                                                                    updateForCAMS_NonPE(ArticleID, sub,dueDate);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                //uploadZIPToFTP("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", dPath + "\\CAMS Delivery\\FROM CUP\\" + ftpFileName, "//from_cams//" + ftpFileName, "none");
                                                            }

                                                        }
                                                        else
                                                        {
                                                            downloadZip("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", dPath + "\\CAMS Delivery\\FROM CUP\\" + ftpFileName, "//from_cams//" + ftpFileName, "none",dPath);
                                                            Console.WriteLine("CAMS:" + sub);
                                                        }
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        Console.WriteLine(e.Message);
                                                    }

                                                    if (File.Exists(dPath + "\\CAMS Delivery\\FROM CUP\\" + ftpFileName))
                                                    {
                                                        try
                                                        {
                                                            bool result = moveProcessedFile("camsftp.cambridge.org", "21", "luminad", "(n>2Uwc;?ZA`pqJ:", "/from_cams/Processed//" + ftpFileName, "/from_cams/" + ftpFileName, "none");
                                                            //FileInfo f = new FileInfo(dPath + "\\CAMS Delivery\\FROM CUP\\" + ftpFileName);
                                                            //f.CopyTo(dPath + "\\CAMS Delivery\\TO CUP\\" + ftpFileName);
                                                            string dbQuery = "Update ldl_trn_articletracker set CUP_PII='" + PII + "' where foldercode='" + ArticleID + "'";
                                                            UpdateQuery(dbQuery);
                                                            UpdateQuery("Update ldl_trn_articletracker set electronicDeliverablesReceived=getdate() where foldercode='" + ArticleID + "' and electronicDeliverablesReceived is null");

                                                            if (dueDate.Trim() == "")
                                                            {
                                                                dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                                                                dueDate = dueDate.Replace("Due Date : ", "");
                                                            }

                                                            if (!sub.ToLower().Contains("re-supply"))
                                                            {
                                                                try
                                                                {

                                                                    copyFinalPDFToOnlinePDF(dPath);

                                                                    if (dt1.Rows.Count == 1)
                                                                    {
                                                                        DataRow dr1 = dt1.Rows[0];
                                                                        string stageID = dr1[0].ToString();
                                                                        if (stageID == "STM1012")
                                                                        {                                                                            
                                                                            //UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com','capssupport@luminad.com','Article is in author stage still and CAMS request received: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')");
                                                                            string tatDays = getTatDays(dr1[1].ToString(), "STM1033");
                                                                            updateForSupply(ArticleID, dueDate, "STM1012", tatDays,"STM1033","","No");
                                                                            UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com;cup.xmlpdy@luminad.com','capssupport@luminad.com','CAMS request received in AUTHOR stage: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + "<p>From: " + fromMailID + "</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')");
                                                                            createLog(dPath, "CAMS Package Downloaded. As this article is in Author Queue so not send for OQC and moved to First View.");

                                                                            //rishi
                                                                            //updateForCAMSReceived(ArticleID, dPath, dr1[1].ToString(), dr1[2].ToString());
                                                                        }
                                                                        if (stageID == "STM1032")
                                                                        {
                                                                            string tatDays = getTatDays(dr1[1].ToString(),"STM1033");
                                                                            
                                                                            //updateForSupply(ArticleID, dueDate, "STM1032", ".5", "STM1033", "");
                                                                            //UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com','capssupport@luminad.com','CAMS request received in PE stage: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team and artilce move to First View stage.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')");

                                                                            //STM1082 // For First View Waiting \\  For Resupply STM1033
                                                                            //ArtMetaInfo.xml

                                                                            string artMetaInfo = dirStartPath + ArticleID;
                                                                            string[] artMetaFile = Directory.GetFiles(artMetaInfo, "*-ArtMetaInfo.xml", SearchOption.TopDirectoryOnly);
                                                                            string[] epsFiles = Directory.GetFiles(artMetaInfo, "*.eps", SearchOption.TopDirectoryOnly);

                                                                            if (epsFiles.Length > 0 && artMetaFile.Length == 0)
                                                                            {
                                                                                updateForSupply(ArticleID, dueDate, "STM1032", ".5", "STM1033", " and no ArtMetaInfo available","Yes");
                                                                                UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com;cup.xmlpdy@luminad.com','capssupport@luminad.com','CAMS request received in PE stage: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team and artilce move to First View stage as no ArtMetaInfo avaialable but eps file avaialble.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')");
                                                                                createLog(dPath, "CAMS Package Downloaded. Move to First View as no ArtMetaInfo xml avaialable.");
                                                                                Console.WriteLine("Move to FV as no ArtMetaInfo xml avaialable.");
                                                                            }                                                                            
                                                                            else
                                                                            {
                                                                                try
                                                                                {
                                                                                    if (dr1[3].ToString().ToLower() == "tex")
                                                                                    {
                                                                                        updateForSupply(ArticleID, dueDate, "STM1032", ".5", "STM1033", " tex artilce moved to FV", "YES");
                                                                                        UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com;cup.latex@luminad.com;cup.xmlpdy@luminad.com','capssupport@luminad.com','CAMS request for LaTeX received in PE stage: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')");
                                                                                        updateForCAMSReceived(ArticleID, dPath, dr1[1].ToString(), dr1[2].ToString(),sub);
                                                                                        createLog(dPath, "CAMS Package Downloaded. And artilce moved to FV as it is TeX artilce.");
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        updateForSupply(ArticleID, dueDate, "STM1032", ".5", "STM1082", "", "YES");
                                                                                        UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com;cup.xmlpdy@luminad.com','capssupport@luminad.com','CAMS request received in PE stage: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')");
                                                                                        updateForCAMSReceived(ArticleID, dPath, dr1[1].ToString(), dr1[2].ToString(),sub);
                                                                                        createLog(dPath, "CAMS Package Downloaded. And request send for OQC bundling.");
                                                                                    }
                                                                                }
                                                                                catch(Exception e) {
                                                                                    Console.WriteLine(e.Message);
                                                                                    updateForSupply(ArticleID, dueDate, "STM1032", ".5", "STM1082", "", "YES");
                                                                                    UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com;cup.xmlpdy@luminad.com','capssupport@luminad.com','CAMS request received in PE stage: " + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')");
                                                                                    updateForCAMSReceived(ArticleID, dPath, dr1[1].ToString(), dr1[2].ToString(),sub);
                                                                                    createLog(dPath, "CAMS Package Downloaded. And request send for OQC bundling.");
                                                                                }
                                                                                //rishi                                                                                
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                catch (Exception e) {
                                                                    Console.WriteLine(e.Message);
                                                                }
                                                                // LDL-CUP-JID-AID                                                                
                                                            }
                                                            else
                                                            {
                                                                //string[] articleDetails = ArticleID.Split('-');
                                                                if (dt1.Rows.Count == 1)
                                                                {
                                                                    DataRow dr1 = dt1.Rows[0];
                                                                    string tatDays = getTatDays(dr1[1].ToString(), "STM1079");
                                                                    updateForReSupply(ArticleID, dueDate, tatDays);
                                                                }
                                                                
                                                            }
                                                        }
                                                        catch { }
                                                        PII = "";
                                                    }
                                                    else
                                                    {
                                                        if (dt1.Rows.Count == 1)
                                                        {
                                                            DataRow dr1 = dt1.Rows[0];
                                                            string stageID = dr1[0].ToString();
                                                            if (stageID == "STM1012" || stageID == "STM1032")
                                                            {
                                                                try
                                                                {
                                                                    Console.WriteLine("Mail:" + sub);
                                                                }
                                                                catch { }
                                                                continue;
                                                            }
                                                        }
                                                    }
                                                }
                                            }                                            
                                        }

                                        // Update when mail received for Package Import Complete
                                        if (!path.ToLower().Contains("notfound") && moveToRP == true && !fromMailID.ToLower().Contains("lumina") && path.Contains("LDL-CUP-"))
                                        {
                                            string[] articleDetails = ArticleID.Split('-');
                                            string aid = "";
                                            if (articleDetails[2] == "ANM" && articleDetails.Length == 5)
                                            {
                                                aid = "-" + articleDetails[4];
                                            }
                                            string moveTo="STM1049";

                                            if(msgSubject.Contains("_AM_")){
                                                moveTo="STM1054";
                                            }

                                            try
                                            {
                                                if ((articleDetails[2] == "BJN" || articleDetails[2] == "NEU" || articleDetails[2] == "WET" || articleDetails[2] == "WSC" || articleDetails[2] == "CJN" || articleDetails[2] == "CTS" || articleDetails[2] == "INP") && moveTo == "STM1049")
                                                {
                                                    string query1 = "select count(*) from ldl_trn_CUP_CAMSDetails c where c.PublisherID='P1031' and c.DeliveryType='Supply' and c.JournalID='" + articleDetails[2] + "' and c.ArticleID='" + articleDetails[3] + "'";
                                                    DataTable dt1 = GetDataFromDB(query1);
                                                    if (dt1.Rows.Count == 1)
                                                    {
                                                        DataRow dr1 = dt1.Rows[0];
                                                        string rowEntryCount = dr1[0].ToString();
                                                        if (rowEntryCount == "0")
                                                        {
                                                            moveTo = "STM1054";
                                                        }
                                                    }
                                                }
                                            }
                                            catch { }

                                            if (moveTo == "STM1049")
                                            {
                                                string dbQuery = "Update ldl_trn_articletracker set stageid='" + moveTo + "' where foldercode='" + ArticleID + "';";
                                                UpdateQuery( " Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " + "stageid,starttime,status,remark,userid)  values ('P1031','" + articleDetails[2] + "','" + articleDetails[3] + aid + "'" + ",'STM1049',getdate(),'STAGE-CHANGE','CAMS Package Import Complete','caps1')");
                                                UpdateQuery(dbQuery);

                                                UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1031','cup.caps@luminad.com','rishiraj.sharma@luminad.com;sandeep.kumar@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Article moved to Ready for Print.</p><br/>" + mb + "',getdate())");
                                            }
                                            else
                                            {
                                                UpdateQuery("Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " + "stageid,starttime,status,remark,userid)  values ('P1031','" + articleDetails[2] + "','" + articleDetails[3] + aid + "'" + ",'STM1085',getdate(),'STAGE-CHANGE','Mail received for AM CAMS Package Import Complete','caps1')");
                                                UpdateQuery("Update ldl_trn_parallelprocess_mapping set stageid='" + moveTo + "' where foldercode='" + ArticleID + "'");

                                                UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1031','cup.caps@luminad.com','rishiraj.sharma@luminad.com;sandeep.kumar@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Article moved to AM Complete.</p><br/>"  + mb + "',getdate())");

                                            }
                                            
                                        }

                                        if (publisher == "P1031" && !sub.Contains("CAMS Content Delivery") && !fromMailID.ToLower().Contains("lumina") && !recFrom.ToLower().EndsWith(".caps") && !fromMailID.Contains("cup.journals") && !fromMailID.Contains("cupproduction"))
                                        {
                                            string sendTo = "cup.pm@luminad.com";
                                            if (msgSubject.Contains("REC") || msgSubject.Contains("INS") || mailbody.Contains("IPG") || mailbody.Contains("EAG") || mailbody.Contains("MSC") || mailbody.Contains("NWS") || mailbody.Contains("MDY") || mailbody.Contains("TLP") || mailbody.Contains("ASB") || mailbody.Contains("GMJ") || mailbody.Contains("XPS") || mailbody.Contains("AER") || mailbody.Contains("CTY"))
                                            {
                                                //sendTo = ";cup.pm@luminad.com";
                                            }
                                            string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','" + sendTo + "','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + "<p>From: " + fromMailAddress + " ("+fromMailID+")</p><p>To: "+toMailID+"</p><p>CC: "+toCC+"</p><p>" + recTime + "</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')";

                                            if (!path.Contains("NotFound\\") && (path.Contains("LDL-CUP-CTY-") || path.Contains("LDL-CUP-XPS-") || path.Contains("LDL-CUP-ASB-")))
                                            {
                                                string authorPDFPath = dirStartPath + ArticleID;
                                                string[] dirFiles = Directory.GetFiles(authorPDFPath, "*_Author_*.pdf");
                                                if (dirFiles.Length == 0)
                                                {
                                                    dirFiles = Directory.GetFiles(authorPDFPath, "*_AUTHOR_*.pdf");
                                                }
                                                if (dirFiles.Length > 0)
                                                {
                                                    DataTable dt1 = GetDataFromDB("select t.stageid,d.email,t.publisherid from ldl_trn_articletracker t left join ldl_trn_authorDetails d on t.ARTICLEID=d.articleID and t.JOURNALID=d.journalid where d.IsCorrespondingAuthor='1' and foldercode='" + ArticleID + "'");
                                                    if (dt1.Rows.Count == 1)
                                                    {
                                                        DataRow dr1 = dt1.Rows[0];
                                                        string stageID = dr1[0].ToString();
                                                        string email = dr1[1].ToString();
                                                        //if (stageID == "STM1012" && fromMailID.ToLower().Contains(email.ToLower()))
                                                        if (stageID == "STM1012" && path.Contains("LDL-CUP-CTY"))
                                                        {
                                                            dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com,ctyproduction@cambridge.org','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')";
                                                        }
                                                        else if (stageID == "STM1012" && path.Contains("LDL-CUP-XPS"))
                                                        {
                                                            dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com,axrahman@cambridge.org','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')";
                                                        }
                                                        else if (stageID == "STM1012" && path.Contains("LDL-CUP-ASB"))
                                                        {
                                                            dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime,AttachmentLocation) values ('P1031','cup.caps@luminad.com','cup.pm@luminad.com,asbproduction@cambridge.org','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate(),'" + attachmentToForward + "')";
                                                        }
                                                    }
                                                }
                                            }

                                            UpdateQuery(dbQuery);                                            
                                        }


                                        // For CUP First Proof- Correction mail received on cup.journals
                                        //if (!path.Contains("NotFound\\") && path.Contains("LDL-CUP-") && sub.Contains("First Proof-") && PII.Trim() == "" && (mailId.ToLower().Contains("cup.journals@luminad.com") || mailId.ToLower().Contains("cup.caps@luminad.com")))
                                        //&& sub.Contains("First Proof-") && PII.Trim() == ""
                                        if (!path.Contains("NotFound\\") && !fromMailID.ToLower().Contains("lumina") && (path.Contains("LDL-CUP-POL-") || path.Contains("LDL-CUP-BME-") || path.Contains("LDL-CUP-DOH-") || path.Contains("LDL-CUP-INS-") || path.Contains("LDL-CUP-REC-") || path.Contains("LDL-CUP-AER-") || path.Contains("LDL-CUP-CTY-") || path.Contains("LDL-CUP-XPS-") || path.Contains("LDL-CUP-ASB-") || path.Contains("LDL-CUP-NWS-") || path.Contains("LDL-CUP-REC-") || path.Contains("LDL-CUP-EAG-") || path.Contains("LDL-CUP-MDY-") || path.Contains("LDL-CUP-PAS-") || path.Contains("LDL-CUP-IPG-") || path.Contains("LDL-CUP-GMJ-") || path.Contains("LDL-CUP-EAG-") || path.Contains("LDL-CUP-MSC-") || path.Contains("LDL-CUP-NWS-") || path.Contains("LDL-CUP-TLP-")) && !recFrom.ToLower().EndsWith(".caps") && (mailId.ToLower().Contains("cup.journals@luminad.com") || mailId.ToLower().Contains("cupproduction@luminad.com") || mailId.ToLower().Contains("cup.caps@luminad.com")))                                        
                                        {
                                            bool isReqUpdate = false;
                                            foreach (MessagePart mpart in message.FindAllAttachments())
                                            {
                                                string fileNameAttach = mpart.FileName;
                                                string file_name_attach = mpart.ContentType.MediaType;

                                                if (fileNameAttach.Contains("winmail"))
                                                    continue;
                                                

                                                if (publisher == "P1031" && (sub.Contains("Proof-") || sub.Contains("First Proof-")) && fileNameAttach.ToLower().Contains(".pdf"))
                                                {
                                                    if (!Directory.Exists(dirStartPath + ArticleID + "\\correction\\"))
                                                    {
                                                        Directory.CreateDirectory(dirStartPath + ArticleID + "\\correction\\");
                                                    }

                                                    if (fileNameAttach.Replace("#", "").Length > 70)
                                                    {
                                                        int lent = fileNameAttach.Length;
                                                        fileNameAttach = fileNameAttach.Replace("#", "").Substring(0, 57) + "_" + fileNameAttach.Replace("#", "").Substring(lent - 4);
                                                    }

                                                    if (!File.Exists(dirStartPath + ArticleID + "\\correction\\" + fileNameAttach))
                                                    {
                                                        try
                                                        {                                                           
                                                            FileInfo finfoo = new FileInfo(dirStartPath + ArticleID + "\\correction\\" + fileNameAttach.Replace("#", ""));
                                                            mpart.Save(finfoo);
                                                        }catch(Exception e){
                                                            Console.WriteLine(e.Message);
                                                        }
                                                        isReqUpdate = true;
                                                    }
                                                }
                                            }

                                            //Console.WriteLine(sub);
                                            if (!recFrom.ToLower().Contains("caps") && !recFrom.ToLower().Contains("lumina") && isReqUpdate == false && (sub.Contains("PLEASE DO NOT RESPOND DIRECTLY TO THIS EMAIL") || sub.ToLower().Contains("reminder")))
                                            {
                                                isReqUpdate = true;
                                            }

                                            if (isReqUpdate == false && !recFrom.ToLower().Contains("caps")  && !recFrom.ToLower().Contains("lumina") && !path.Contains("NotFound\\"))
                                            {
                                                string authorPDFPath = dirStartPath + ArticleID;
                                                 string[] dirFiles = Directory.GetFiles(authorPDFPath, "*_Author_*.pdf");
                                                 if (dirFiles.Length > 0)
                                                 {
                                                     isReqUpdate = true;
                                                 }
                                            }

                                            if (mailbody.ToLower().Contains("out of office") || mailbody.ToLower().Contains("out-of-office") || sub.ToLower().Contains("automatic reply"))
                                            {
                                                isReqUpdate = false;
                                            }

                                            if (isReqUpdate == true)
                                            {
                                                if (!Directory.Exists(dirStartPath + ArticleID + "\\correction\\"))
                                                {
                                                    Directory.CreateDirectory(dirStartPath + ArticleID + "\\correction\\");
                                                }

                                                bool result =updationForCUPFirstProof(publisher,ArticleID);
                                                
                                                try
                                                {
                                                    if (result == true)
                                                    {

                                                        // run exe to create xml to html for CUP
                                                        try
                                                        {
                                                            createXMLToHTML(ArticleID, dirStartPath);                                                            
                                                        }
                                                        catch { }


                                                        string mailNewPath = dirStartPath + ArticleID + "\\correction\\" + fileToSave + ".mail";
                                                        StreamWriter sw1 = new StreamWriter(mailNewPath, true, Encoding.ASCII);
                                                        if (longsubjectBool == true)
                                                        {
                                                            sw1.WriteLine("Subject: " + longsubject + "\n");
                                                        }
                                                        sw1.WriteLine("From: " + fromMailID);
                                                        sw1.WriteLine("To  : " + toMailID);
                                                        if (toCC != "" && toCC != fromMailID)
                                                        {
                                                            sw1.WriteLine("CC  : " + toCC);
                                                        }
                                                        sw1.WriteLine("");
                                                        sw1.WriteLine(mailbody);
                                                        sw1.Close();                                                    
                                                    }
                                                }
                                                catch (Exception e)
                                                {
                                                    Console.WriteLine(e.Message);
                                                }
                                            }
                                        }

                                        try {
                                            sw = new StreamWriter(path, true, Encoding.ASCII);
                                        }
                                        catch(Exception e)
                                        {
                                            Console.WriteLine(e.Message);
                                            //13Sep2017_ldps_10.06.44_ldps_	--changjun 自动回复 Your page proofs have been submitted for article AO-296089
                                            String[] ar = fileToSave.Split('_');
                                            String nname = "";
                                            for (int ar1 = 0; ar1 < ar.Length - 1; ar1++)
                                            {
                                                nname = nname + "_" + ar[ar1];
                                            }

                                            nname = nname + "_Subject missing";
                                            nname = nname.Trim('_');

                                            path = path.Replace(fileToSave, nname);
                                            //sw = new StreamWriter(path, true, Encoding.UTF8);
                                            sw = new StreamWriter(path, true, Encoding.ASCII);
                                            longsubjectBool = true;
                                        }
                                        

                                        if (!path.Contains("NotFound\\") && path.Contains("LDL-OSA-") && !recFrom.ToLower().EndsWith(".caps"))
                                        {
                                            string authorPDFPath = dirStartPath + ArticleID;

                                            string[] dirFiles = Directory.GetFiles(authorPDFPath, "*_Author_*.pdf");

                                            bool folderEXist = Directory.Exists(path);

                                            if (dirFiles.Length > 0)
                                            {
                                                DataTable dt1 = GetDataFromDB("select t.stageid,d.email,t.publisherid from ldl_trn_articletracker t left join ldl_trn_authorDetails d on t.ARTICLEID=d.articleID and t.JOURNALID=d.journalid where d.IsCorrespondingAuthor='1' and foldercode='"+ArticleID+"'");                                                
                                                if (dt1.Rows.Count == 1)
                                                {
                                                    DataRow dr1 = dt1.Rows[0];
                                                    string stageID = dr1[0].ToString();
                                                    string email = dr1[1].ToString();
                                                    string pid = dr1[2].ToString();
                                                    //if ((stageID == "STM1048" || stageID == "STM1049" ) && email!="" && fromMailID.Contains(email))
                                                    if (email != "" && fromMailID.Contains(email))
                                                    {
                                                        string sendMailTO = "capssupport@luminad.com";
                                                        string sendMailCC = "osa-manager@luminad.com,capssupport@luminad.com";
                                                        if (path.Contains("LDL-OSA-AO-"))
                                                        {
                                                            sendMailTO = "aocorrections@osa.org";
                                                        }
                                                        else if (path.Contains("LDL-OSA-AOP-"))
                                                        {
                                                            sendMailTO = "aopcorrections@osa.org";
                                                        }
                                                        else if (path.Contains("LDL-OSA-JA-"))
                                                        {
                                                            sendMailTO = "josaacorrections@osa.org";
                                                        }
                                                        else if (path.Contains("LDL-OSA-JB-"))
                                                        {
                                                            sendMailTO = "josabcorrections@osa.org";
                                                        }
                                                        else if (path.Contains("LDL-OSA-JOCN-"))
                                                        {
                                                            sendMailTO = "jocncorrections@osa.org";
                                                        }
                                                        else if (path.Contains("LDL-OSA-OL-"))
                                                        {
                                                            sendMailTO = "olcorrections@osa.org";
                                                        }
                                                        else if (path.Contains("LDL-OSA-OPTICA-"))
                                                        {
                                                            sendMailTO = "opticacorrections@osa.org";
                                                        }
                                                        //else if (path.Contains("LDL-OSA-PR-"))
                                                        //{
                                                        //    sendMailTO = "slliu@siom.ac.cn,osa-manager@luminad.com,capssupport@luminad.com";
                                                        //}

                                                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('" + pid + "','" + mailId + "','" + sendMailTO + "','" + sendMailCC + "','" + msgSubject.Replace("'","''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate())";
                                                        UpdateQuery(dbQuery);

                                                        //SendEmail_WithToMailID(msgSubject, "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb, mailId, mailPassword, ArticleID, sendMailTO,sendMailCC);
                                                        //osa-manager@luminad.com
                                                    }
                                                }
                                            }
                                        }

                                        // For AOAC and ASTM

                                        if (!path.Contains("NotFound\\") && (path.Contains("LDL-ASTM-") || path.Contains("LDL-AOAC-")) && !recFrom.ToLower().EndsWith(".caps"))
                                        {
                                            string authorPDFPath = dirStartPath + ArticleID;
                                            string[] dirFiles = Directory.GetFiles(authorPDFPath, "*_Author_*.pdf");

                                           bool folderEXist = Directory.Exists(path);

                                            if (dirFiles.Length > 0)
                                            {                                               
                                                DataTable dt1 = GetDataFromDB("select t.stageid,d.email,t.publisherid from ldl_trn_articletracker t left join ldl_trn_authorDetails d on t.ARTICLEID=d.articleID and t.JOURNALID=d.journalid where d.IsCorrespondingAuthor='1' and foldercode='" + ArticleID + "'");

                                                if (dt1.Rows.Count == 1)
                                                {
                                                    DataRow dr1 = dt1.Rows[0];
                                                    string stageID = dr1[0].ToString();
                                                    string email = dr1[1].ToString();
                                                    string pid = dr1[2].ToString();
                                                    //if ((stageID == "STM1048" || stageID == "STM1049" ) && email!="" && fromMailID.Contains(email))
                                                    if (email != "")
                                                    {
                                                        string sendMailTO = "capssupport@luminad.com";
                                                        string sendMailCC = "capssupport@luminad.com";
                                                        if (path.Contains("LDL-ASTM-"))
                                                        {
                                                            sendMailTO = "astm.manager@luminad.com";
                                                        }
                                                        else if (path.Contains("LDL-AOAC-"))
                                                        {
                                                            sendMailTO = "aoac.manager@luminad.com";
                                                        }

                                                        //SendEmail_WithToMailID(msgSubject, "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb, mailId, mailPassword, ArticleID, sendMailTO, sendMailCC);
                                                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('" + pid + "','" + mailId + "','" + sendMailTO + "','" + sendMailCC + "','" + msgSubject.Replace("'", "''") + "','" + "<p>This mail is forwarded by CAPS Team.</p><br/>" + mb + "',getdate())";
                                                        UpdateQuery(dbQuery);
                                                        //osa-manager@luminad.com
                                                    }
                                                }
                                            }
                                        }

                                        if (path.Contains("LDL-HK-") && (msgSubject.ToUpper().Replace(".", "-").Trim() == ArticleID.Replace("LDL-HK-", "").ToUpper() 
                                            + " ERRATA" || msgSubject.ToUpper().Replace(".", "-").Trim() == "ERRATA " + ArticleID.Replace("LDL-HK-", "").ToUpper()||
                                            msgSubject.ToUpper().Replace(".", "-").Trim() == ArticleID.Replace("LDL-HK-", "").ToUpper()
                                            + " ERRATUM" || msgSubject.ToUpper().Replace(".", "-").Trim() == "ERRATUM " + ArticleID.Replace("LDL-HK-", "").ToUpper()))
                                        {
                                            if (msgSubject.ToLower().Contains("errat") && path.Contains("LDL-HK-"))
                                            {
                                                UpdateQuery("Update ldl_trn_Articletracker set stageid='STM1033' where foldercode='" + ArticleID + "'");

                                                DataTable dt1 = GetDataFromDB("Select publisherid,journalid,articleid from ldl_trn_articletracker where FOLDERCODE='" + ArticleID + "'");
                                                if (dt1.Rows.Count == 1)
                                                {
                                                    DataRow dr1 = dt1.Rows[0];
                                                    string publisherid = dr1[0].ToString();
                                                    string journalID = dr1[1].ToString();
                                                    string articleid = dr1[2].ToString();

                                                    UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister (publisherid,journalid,articleid, " +
                                                "stageid,starttime,status,remark,userid)  values ('" + publisherid + "','" + journalID + "','" + articleid + "'" +
                                                ",'STM1033',getdate(),'STAGE-CHANGE','Article moved to ADD FOLIOS as mail received.','caps1')");
                                                }
                                            }
                                        }

                                        // For HK
                                        if (path.Contains("LDL-HK-") && !recFrom.ToLower().EndsWith(".caps"))
                                        {
                                            string authorPDFPath = dirStartPath + ArticleID;
                                            string[] dirFiles = Directory.GetFiles(authorPDFPath, "*_Author_*.pdf");
                                            bool folderEXist = Directory.Exists(path);

                                            if (dirFiles.Length > 0)
                                            {
                                                //DataTable dt1 = GetDataFromDB("Select stageid from ldl_trn_articletracker where FOLDERCODE='" + ArticleID + "'");
                                                DataTable dt1 = GetDataFromDB("Select t.stageid,sm.STAGENAME from ldl_trn_articletracker t left join ldl_mst_StageMaster sm on sm.PublisherID=t.PUBLISHERID and sm.STAGEMASTERID=t.STAGEID where t.FOLDERCODE='" + ArticleID + "'");
                                                if (dt1.Rows.Count == 1)
                                                {
                                                    DataRow dr1 = dt1.Rows[0];
                                                    string stageID = dr1[0].ToString();
                                                    string stageName = dr1[1].ToString();
                                                    if (stageID == "STM1048" || stageID == "STM1049" || stageID == "STM1035" || stageID == "STM1033" || stageID == "STM1032" || stageID == "STM1017" || stageID == "STM1017" || stageID == "STM1016" || stageID == "STM1013")
                                                    {
                                                        //SendEmail(msgSubject, mailbody, mailId, mailPassword, ArticleID);                                                   
                                                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1027','" + mailId + "','hk.manager@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>We have received late corrections from the author when the article is in <b>" + stageName + "</b> stage.</p><br/>" + mb + "',getdate())";
                                                        UpdateQuery(dbQuery);
                                                    }
                                                }
                                            }
                                        }

                                        //Send Article mail to Melissa in case of Complete and Ready for Print
                                        if (ArticleID.Contains("ASCE"))
                                        {
                                            DataTable dt1 = GetDataFromDB("Select stageid from ldl_trn_articletracker where FOLDERCODE='" + ArticleID + "'");
                                            if (dt1.Rows.Count == 1)
                                            {
                                                DataRow dr1 = dt1.Rows[0];
                                                string stageID = dr1[0].ToString();

                                                if (stageID == "STM1048" || stageID == "STM1049")
                                                {
                                                    //SendEmail(msgSubject, mailbody, mailId, mailPassword, ArticleID);
                                                    Console.WriteLine("--");
                                                    Console.WriteLine("Path::" + path);
                                                    string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1003','" + mailId + "','asce-manager@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + mb + "',getdate())";
                                                    UpdateQuery(dbQuery);
                                                }
                                            }
                                        }

                                        if (ArticleID.Contains("SPIE"))
                                        {
                                            DataTable dt1 = GetDataFromDB("Select stageid from ldl_trn_articletracker where FOLDERCODE='" + ArticleID + "'");
                                            if (dt1.Rows.Count == 1)
                                            {
                                                DataRow dr1 = dt1.Rows[0];
                                                string stageID = dr1[0].ToString();
                                                if (stageID == "STM1048" || stageID == "STM1049" || stageID == "STM1035" || stageID == "STM1033")
                                                {
                                                    //SendEmail(msgSubject, mailbody, mailId, mailPassword, ArticleID);

                                                    string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1007','" + mailId + "','Melissa.metz@luminad.com,spie-manager@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + mb + "',getdate())";
                                                    UpdateQuery(dbQuery);
                                                }
                                            }
                                        }

                                        string s = mailmsg.Body;

                                        // Send this mail to Melissa  that mail contain subject as "please disregard the ftp transmittal to comp task assignment for"
                                        if (msgSubject.ToLower().Contains("please disregard the ftp transmittal to comp task assignment for"))
                                        {
                                            //SendEmail(msgSubject, mailbody, mailId, mailPassword, ArticleID);
                                            string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1003','" + mailId + "','Melissa.metz@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + mb + "',getdate())";
                                            UpdateQuery(dbQuery);
                                        }
                                        // End Here
                                    }
                                    catch(Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                        //13Sep2017_ldps_10.06.44_ldps_	--changjun 自动回复 Your page proofs have been submitted for article AO-296089
                                        String[] ar = fileToSave.Split('_');
                                        String nname = "";
                                        for (int ar1 = 0; ar1 < ar.Length-1; ar1++)
                                        {
                                            nname = nname + "_" + ar[ar1];
                                        }

                                        nname = nname + "_Subject missing";
                                        nname= nname.Trim('_');

                                        path = path.Replace(fileToSave, nname);
                                        //sw = new StreamWriter(path, true, Encoding.UTF8);
                                        sw = new StreamWriter(path, true, Encoding.ASCII);
                                        longsubjectBool = true;
                                    }

                                    //using (sw = File.CreateText(path))
                                    //{
                                    if (longsubjectBool == true)
                                    {
                                        sw.WriteLine("Subject: " + longsubject + "\n");
                                    }
                                    sw.WriteLine("From: " + fromMailID);
                                    sw.WriteLine("To  : " + toMailID);
                                    if (toCC != "" && toCC != fromMailID)
                                    {
                                        sw.WriteLine("CC  : " + toCC);
                                    }
                                    
                                    sw.WriteLine("");
                                    sw.WriteLine(mailbody);
                                    sw.Close();

                                    if (File.Exists(path))
                                    {
                                        //Console.WriteLine("Mail Path 1:: " + path);
                                        if (!publisher.Contains("P1031"))
                                        {
                                            if (publisher.Contains("P1032"))
                                            {
                                                if (isError == false)
                                                {
                                                    pop3Client.DeleteMessage(i);
                                                }
                                            }
                                            else
                                            {
                                                if (!publisher.Contains("P1030"))
                                                {
                                                    pop3Client.DeleteMessage(i);
                                                }
                                            }
                                        }
                                        if (longsubjectBool == true)
                                        {
                                            pop3Client.DeleteMessage(i);
                                            longsubjectBool = false;
                                        }

                                    }
                                }
                            }
                            else
                            {
                                if (recFrom.ToLower().EndsWith(".caps"))
                                {
                                    string[] files = System.IO.Directory.GetFiles(dirPath, recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + arrMailRecTime[0] + "." + "*_ldps_" + msgSubject + ".mail", System.IO.SearchOption.AllDirectories);
                                    if (files.Length == 0)
                                    {
                                        files = System.IO.Directory.GetFiles(dirPath, "0" + recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + arrMailRecTime[0] + "." + "*_ldps_" + msgSubject + ".mail", System.IO.SearchOption.AllDirectories);
                                    }
                                    if (files.Length > 1)
                                    {
                                        for (int l = 1; l < files.Length; l++)
                                        {
                                            try
                                            {
                                                FileInfo fin = new FileInfo(files[l]);
                                                String newpath = "D:\\CAPSMail\\" + fin.Name;
                                                File.Move(files[l], newpath);
                                            }
                                            catch { }
                                        }
                                    }
                                }

                                if (!publisher.Contains("P1031"))
                                {
                                    if (!publisher.Contains("P1030"))
                                    {
                                        pop3Client.DeleteMessage(i);
                                    }
                                    
                                }
                            }
                            
                            success++;
                        }
                        catch (Exception e)
                        {
                            DefaultLogger.Log.LogError(
                                "TestForm: Message fetching failed: " + e.Message + "\r\n" +
                                "Stack trace:\r\n" +
                                e.StackTrace);
                            fail++;
                        }                       
                        //progressBar.Value = (int)(((double)(count - i) / count) * 100);
                    }
                    
                    if (fail > 0)
                    {
                        //MessageBox.Show(this,
                        //"Since some of the emails were not parsed correctly (exceptions were thrown)\r\n" +
                        //"please consider sending your log file to the developer for fixing.\r\n" +
                        //"If you are able to include any extra information, please do so.",
                        //"Help improve OpenPop!");
                    }
                }
                //pop3Client.Dispose();
                if (pop3Client.Connected)
                {
                    pop3Client.Disconnect();
                }
                Console.WriteLine("Save Mails from server done.");
            }
            catch (InvalidLoginException) { }
            catch (PopServerNotFoundException) { }
            catch (PopServerLockedException){}
            catch (LoginDelayException){}            
            catch(Exception e) {
                Console.WriteLine(e.Message);
            }
            finally
            {
                
            }
        }

        // Update for Mail received as a History article wise
        private void UpdateInArticleMovementRegister(string ArticleID, string currentStage,string remark)
        {
            try
            {
                string[] arrStage = currentStage.Split(':');
                string stageID = "", stageName = "";
                if (arrStage.Length == 2)
                {
                    stageID = arrStage[0];
                    stageName = " in " + arrStage[1];
                }

                string[] articleDetails = ArticleID.Split('-');
                string aid = "";
                if (articleDetails[2] == "ANM" && articleDetails.Length == 5)
                {
                    aid = "-" + articleDetails[4];
                }

                UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " + "stageid,starttime,status,remark,userid)  values ('P1031','" + articleDetails[2] + "','" + articleDetails[3] + aid + "'" + ",'"+stageID+"',getdate(),'Mail Movement','"+remark+stageName+"','caps1')");
            }
            catch { }
        }

        public string getCurrentStageAndName(string publisher, string ArticleID)
        {
            // Retrun Colon seprated StageID:StageName

            string currentStage = "";
            
            try
            {
                if (publisher == "P1031" && ArticleID.Contains("LDL-CUP-"))
                {
                    
                    DataTable currentStageDataTable = GetDataFromDB("Select t.stageid,sm.stagename from ldl_trn_articletracker t left join ldl_mst_stagemaster sm on sm.stagemasterid=t.stageid and sm.publisherid=t.publisherid where t.FOLDERCODE='" + ArticleID + "'");
                    if (currentStageDataTable.Rows.Count == 1)
                    {
                        DataRow currentStageDataRow = currentStageDataTable.Rows[0];
                        currentStage = currentStageDataRow[0].ToString();
                        currentStage = currentStage + ":" + currentStageDataRow[1].ToString();
                    }
                }
            }
            catch { }
            return currentStage;
        }

        public void copyFinalPDFToOnlinePDF(string dPath)
        {
            try
            {                
                string folder = dPath;
                var files = new DirectoryInfo(folder).GetFiles("*_FINALED_*.*");

                string lasUpdatedFile = "";
                DateTime lastupdated = DateTime.MinValue;
                foreach (System.IO.FileInfo f in files)
                {
                    if (f.LastWriteTime > lastupdated)
                    {
                        lastupdated = f.LastWriteTime;
                        lasUpdatedFile = f.FullName;
                    }
                }
                
                if (File.Exists(lasUpdatedFile))
                {
                    string outFile = Regex.Replace(lasUpdatedFile, "_FINALED_[0-9]+", "_ONLINE");
                    File.Copy(lasUpdatedFile, outFile, true);
                }

            }
            catch { }
        }

        private void updateForCAMS_NonPE(string ArticleID, string sub, string dueDate)
        {
            try
            {
                string[] articleDetails = ArticleID.Split('-');
                string aid = "";
                if (ArticleID.Contains("ANM"))
                {
                    try
                    {

                        aid = "-" + articleDetails[4];
                    }
                    catch (Exception e)
                    {
                        aid = "";
                        Console.WriteLine(e.Message);
                    }
                }                
                string dbQuery = "insert into ldl_trn_CUP_CAMSDetails (publisherid,journalid,articleid,downloadtime,deliverytype,CAMSDueDate,CalculatedCAMSDueDate,ActionTaken) values ('P1031','" + articleDetails[2] + "','" + articleDetails[3] + aid + "',getdate(),'Supply','" + dueDate + "',null,'No')";
                UpdateQuery(dbQuery);
            }
            catch { }
        }

        private void updateForAM(string ArticleID, string sub, string dueDate)
        {
            try
            {
                // ACCEPTED MANUSCRIPT 83   ACCEPTED MANUSCRIPT RESUPPLY 84
                //AM_DueDate
                //AM_CurrentStatus

                DataTable dt1 = GetDataFromDB("select t.journalid,t.articleid,t.publisherid from ldl_trn_articletracker t where foldercode='" + ArticleID + "'");
                if (sub.Trim().ToLower().Contains("re-supply"))
                {

                    if (dt1.Rows.Count == 1)
                    {
                        DataRow dr1 = dt1.Rows[0];
                        string jID = dr1[0].ToString();
                        string aID = dr1[1].ToString();
                        string pID = dr1[2].ToString();
                        UpdateQuery("Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " + "stageid,starttime,status,remark,userid)  values ('P1031','" + jID + "','" + aID + "'" + ",'STM1052',getdate(),'STAGE-CHANGE','Mail received for AM CAMS RE-SUPPLY','caps1')");
                        dt1 = GetDataFromDB("select foldercode from ldl_trn_parallelprocess_mapping  where foldercode='" + ArticleID + "'");
                        if (dt1.Rows.Count == 1)
                        {
                            UpdateQuery("Update ldl_trn_parallelprocess_mapping set stageid='STM1052' where foldercode='" + ArticleID + "'");
                        }
                        else
                        {
                            UpdateQuery("Insert into ldl_trn_parallelprocess_mapping (PUBLISHERID,JOURNALID,ARTICLEID,STAGEID,FOLDERCODE,ISALLOCATED) values ('P1031','" + jID + "','" + aID + "','STM1052','" + ArticleID + "','N')");
                        }
                        UpdateQuery("insert into ldl_trn_CUP_AMCAMSDetails (publisherid,journalid,articleid,downloadtime,DeliveryType,AMCAMSDueDate) values ('P1031','" + jID + "','" + aID + "',getdate(),'RE-SUPPLY','" + dueDate + "')");
                        UpdateQuery(" Update ldl_trn_Articletracker set AM_DueDate='" + dueDate + "', AM_CurrentStatus='Re-Supply' where foldercode='"+ArticleID+"'");
                    }
                }
                else
                {
                    if (dt1.Rows.Count == 1)
                    {
                        DataRow dr1 = dt1.Rows[0];
                        string jID = dr1[0].ToString();
                        string aID = dr1[1].ToString();
                        string pID = dr1[2].ToString();
                        UpdateQuery("Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " + "stageid,starttime,status,remark,userid)  values ('P1031','" + jID + "','" + aID + "'" + ",'STM1055',getdate(),'STAGE-CHANGE','Mail received for AM CAMS SUPPLY','caps1')");
                        dt1 = GetDataFromDB("select foldercode from ldl_trn_parallelprocess_mapping  where foldercode='" + ArticleID + "'");
                        if (dt1.Rows.Count == 1)
                        {
                            UpdateQuery("Update ldl_trn_parallelprocess_mapping set stageid='STM1052' where foldercode='" + ArticleID + "'");
                        }
                        else
                        {
                            UpdateQuery("Insert into ldl_trn_parallelprocess_mapping (PUBLISHERID,JOURNALID,ARTICLEID,STAGEID,FOLDERCODE,ISALLOCATED) values ('P1031','" + jID + "','" + aID + "','STM1055','" + ArticleID + "','N')");
                        }
                        UpdateQuery("insert into ldl_trn_CUP_AMCAMSDetails (publisherid,journalid,articleid,downloadtime,DeliveryType,AMCAMSDueDate) values ('P1031','" + jID + "','" + aID + "',getdate(),'SUPPLY','" + dueDate + "')");
                        UpdateQuery(" Update ldl_trn_Articletracker set AM_DueDate='" + dueDate + "', AM_CurrentStatus='Supply' where foldercode='" + ArticleID + "'");
                    }
                }
            }
            catch { }
        }

        private string getArticleIDForAM(string msgSubject, string mailbody, string fileName, string mailId,string sub,Dictionary<string, string> map)
        {
            string ArticleID = "";
            try
            {
                ArticleID = getArticleID1(msgSubject, mailbody, fileName, mailId);
                ArticleID = ArticleID.Replace("PII: ", "");
                //Journal Mnemonic: GEO
                string jid = getDataByRegex(mailbody, "Journal Mnemonic:\\s[A-Z]+");
                jid = jid.Replace("Journal Mnemonic: ", "");
                string PII = ArticleID.Trim();
                ArticleID = PII.Substring(9, 7);
                if (jid.Trim() != "")
                {
                    ArticleID = jid.Trim() + "-" + ArticleID;
                }
                string ftpFileName = getDataByRegex(mailbody, "Package: .*?zip");

                ftpFileName = getDataByRegex(ftpFileName, "cams_.*?zip");
                //continue;
                string dueDate = "";
                dueDate = getDataByRegex(mailbody, "Due Date : .*?20[0-9]{2}");
                dueDate = dueDate.Replace("Due Date : ", "");

                if (sub.Contains("ANM") && jid.Trim() == "ANM")
                {
                    string aid = "";                    
                    if (map.TryGetValue(ArticleID, out aid))
                    {
                        if (aid != "")
                            ArticleID = "ANM-" + aid;
                    }                    
                }
            }
            catch { }
            return ArticleID;
        }

        private string getTatDays(string jid,string stageID)
        {
            string tatDays = "2";
            if (stageID == "STM1079")
            {
                tatDays = "1";
            }
            try
            {
                DataTable dtTATDays = GetDataFromDB("select TATDaysFresh from ldl_mst_StageReportMapping1 p where p.PublisherID='P1031' and p.JournalID='" + jid + "' and p.CAPSQueueStageID='"+stageID+"'");
                
                if (dtTATDays.Rows.Count == 1)
                {
                    DataRow tatRow = dtTATDays.Rows[0];
                    tatDays = tatRow[0].ToString();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return tatDays;
        }

        private string GetFolderCode(string dirStartPath, string foldToSearchIn, string ArticleID)
        {
            
            try
            {
                string[] subdirectoryEntries = Directory.GetDirectories(dirStartPath);
                for (int l = 0; l < subdirectoryEntries.Length; l++)
                {
                    if (foldToSearchIn == "")
                    {
                        if (subdirectoryEntries[l].Contains(ArticleID))
                        {
                            DirectoryInfo di = new DirectoryInfo(subdirectoryEntries[l]);
                            ArticleID = di.Name;
                            break;
                        }
                    }
                    else
                    {
                        if (subdirectoryEntries[l].Contains(ArticleID) && subdirectoryEntries[l].ToLower().Contains(foldToSearchIn.ToLower()))
                        {
                            DirectoryInfo di = new DirectoryInfo(subdirectoryEntries[l]);
                            ArticleID = di.Name;
                            break;
                        }
                        if (subdirectoryEntries[l].Contains(ArticleID.ToUpper()) && subdirectoryEntries[l].ToLower().Contains(foldToSearchIn.ToLower()))
                        {
                            DirectoryInfo di = new DirectoryInfo(subdirectoryEntries[l]);
                            ArticleID = di.Name;
                            break;
                        }
                    }
                }
            }
            catch { }
            return ArticleID;
        }

        public void updateForCAMSReceived(string ArticleID,string path,string jid,string aid,string sub)
        {
            try
            {
                string currentTime = string.Format("{0:yyyy-MM-dd-hhmmss}", DateTime.Now);             

                File.Copy(path + "\\" + jid + "-" + aid + ".xml", path + "\\" + jid + "-" + aid + "_bk"+currentTime+".xml",true);
                string txt = File.ReadAllText(path+"\\"+jid+"-"+aid+".xml",Encoding.UTF8);
                txt = txt.Replace("pmg:currentqueue=\"format_final\"", "pmg:currentqueue=\"online_publish\" ");

                try
                {
                    string eText = getDataByRegex(txt, "<front>.*?</front>");

                    if(File.Exists(path + "\\" + jid + "-" + aid + "_bk" + currentTime + ".xml") && eText.Trim()!="")
                    {
                        if (jid == "BAJ" || jid == "CBT" || jid == "JFP" || jid == "KER" || jid == "PAS" || jid == "NJG" || jid == "PEN" || jid == "PHC" || jid == "PRP" || jid == "IDM")
                        {                            
                            string oldValue = eText;

                            string vol= getDataByRegex(sub, "_V([0-9]+)_");
                            vol = vol.Replace("_V","").Replace("_","");

                            int eNumber = updateForEArticle(jid, aid,vol);
                            string pasUpdate = "";
                            if (jid == "PAS")
                            {
                                pasUpdate = "0";
                            }
                            if (eNumber != 0)
                            {
                                bool isUpdate = false;
                                string ealloc = getDataByRegex(eText, "<elocation-id>.*?</elocation-id>");
                                if (ealloc.Contains("elocation-id"))
                                {
                                    var regex = new Regex(Regex.Escape(ealloc));
                                    //eText =  eText.Replace(ealloc,"<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id>");
                                    eText = regex.Replace(eText, "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id>", 1);
                                    isUpdate = true;
                                }
                                else if (txt.Contains("</issue>"))
                                {
                                    //eText = eText.Replace("</issue>", "</issue><elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id>");
                                    var regex = new Regex(Regex.Escape("</issue>"));
                                    eText = regex.Replace(eText, "</issue><elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id>", 1);
                                    isUpdate = true;
                                }
                                else if (txt.Contains("</volume>"))
                                {
                                    //eText = eText.Replace("</volume>", "</volume><elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id>");
                                    var regex = new Regex(Regex.Escape("</volume>"));
                                    eText = regex.Replace(eText, "</volume><elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id>", 1);
                                    isUpdate = true;
                                }
                                else if (txt.Contains("<history>"))
                                {
                                    //eText = eText.Replace("<history>", "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id><history>");
                                    var regex = new Regex(Regex.Escape("<history>"));
                                    eText = regex.Replace(eText, "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id><history>", 1);
                                    isUpdate = true;
                                }
                                else if (txt.Contains("<permissions>"))
                                {
                                    //eText = eText.Replace("<permissions>", "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id><permissions>");
                                    var regex = new Regex(Regex.Escape("<permissions>"));
                                    eText = regex.Replace(eText, "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id><permissions>", 1);
                                    isUpdate = true;
                                }
                                else if (txt.Contains("<abstract "))
                                {
                                    //eText = eText.Replace("<abstract ", "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "<abstract ");
                                    var regex = new Regex(Regex.Escape("<abstract "));
                                    eText = regex.Replace(eText, "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id><abstract ", 1);
                                    isUpdate = true;
                                }
                                else if (txt.Contains("<kwd-group>"))
                                {
                                    //eText = eText.Replace("<kwd-group>", "<elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "<kwd-group>");
                                    var regex = new Regex(Regex.Escape("<kwd-group>"));
                                    eText = regex.Replace(eText, "<kwd-group><elocation-id>e" + pasUpdate + Convert.ToString(eNumber) + "</elocation-id><kwd-group>", 1);
                                    isUpdate = true;
                                }
                                if (isUpdate == true)
                                {
                                    string tagToRemove = getDataByRegex(eText, "<fpage>.*?</fpage>");
                                    eText = eText.Replace(tagToRemove, "");

                                    tagToRemove = "";
                                    tagToRemove = getDataByRegex(eText, "<lpage>.*?</lpage>");
                                    eText = eText.Replace(tagToRemove, "");

                                    eText = eText.Replace("<volume>", "<volume seq=\"" + pasUpdate + Convert.ToString(eNumber) + "\">"+ vol);
                                    //eText = eText.Replace("<volume>", "<volume seq=\"" + pasUpdate + Convert.ToString(eNumber) + "\">");
                                    eText = eText.Replace(vol+ vol+"</volume>", vol + "</volume>");
                                    eText = eText.Replace("<issue>e.00</issue>", "").Replace("<issue>00</issue>", "").Replace("<pub-date pub-type=\"generated\">", "<pub-date pub-type=\"ppub\">");

                                    string year = string.Format("{0:yyyy}", DateTime.Now);
                                    if (year.Trim() != "") {
                                        eText = eText.Replace("<year/></pub-date>", "<year>" + year + "</year></pub-date>");
                                    }
                                    

                                    UpdateQuery(" Update ldl_trn_article set  CUPeNumber=" + eNumber + " where journalid='" + jid + "' and articleid='" + aid + "'");
                                    UpdateQuery(" Update ldl_trn_CUPeArticleData set  eNumber=" + eNumber + " where journalid='" + jid + "' ");
                                    //txt = eText;
                                    txt = txt.Replace(oldValue,eText);
                                    //File.WriteAllText(path + "\\" + jid + "-" + aid + "elocation.xml", eText, Encoding.UTF8);
                                    //ldl_trn_CUPeArticleData
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("UpdateForCAMS:" + e.Message);
                }

                File.WriteAllText(path + "\\" + jid + "-" + aid + ".xml", txt, Encoding.UTF8);
                UpdateQuery("Update ldl_trn_articletracker set SmallCorrection= isnull(SmallCorrection,0)+1 where foldercode='" + ArticleID + "'");
                runCUP_PSCreate(path,ArticleID,jid,aid);
                UpdateQuery(" Update ldl_trn_paginationpool set  ISLOCKED='N' where FOLDERCODE='" + ArticleID + "'");

                DataTable dt1 = GetDataFromDB("select foldercode from ldl_trn_paginationpool  where foldercode='" + ArticleID + "'");
                 if (dt1.Rows.Count == 0)
                 {
                     UpdateQuery(" Insert into ldl_trn_paginationpool (FOLDERCODE,STATUS,ISLOCKED,PRIORITY,ARTICLEID,DATETIME) values ('" + ArticleID + "','PROCESSED','N','3','" + aid + "',getdate()) ");
                 }

                try
                {
                    //JMSProducer jProd = new JMSProducer();
                    //jProd.postCAMSMessage(ArticleID, "CUP CAMS REQUEST");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("upateForCAMSReceived: "+ e.Message);
            }
        }

        private int updateForEArticle(string jid, string aid,string vol)
        {
            int eNumber =0;
            try
            {
                string query = "";
                if (vol.Trim() != "")
                {
                    query = "select isnull(eNumber+1,0) from ldl_trn_CUPeArticleData where journalid='" + jid + "' and volume='" + vol+"'";                    
                }
                else
                {
                    query = "select isnull(eNumber+1,0) from ldl_trn_CUPeArticleData where journalid='" + jid + "'";
                }
                DataTable dt1 = GetDataFromDB(query);
                if (dt1.Rows.Count > 0)
                {
                    DataRow dr1 = dt1.Rows[0];
                    eNumber = Convert.ToInt32(dr1[0].ToString());
                }                
                if (eNumber == 0 && vol.Trim() != "")
                {
                    eNumber = 1;
                    UpdateQuery("Update ldl_trn_CUPeArticleData set eNumber=1,volume='" + vol + "' where journalid='" + jid + "'");
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Update for EArticle: "+e.Message);
            }
            return eNumber;
        }

        private void runCUP_PSCreate(string path,string ArticleID,string jid,string aid)
        {
            try
            {
                string arg_3dFile = path + "\\" + jid + "-" + aid + ".3d";
                
                string exePath = ConfigurationManager.AppSettings["PSCreatePath"].ToString();
                ProcessStartInfo info = new ProcessStartInfo(exePath, arg_3dFile);
                info.CreateNoWindow = true;
                info.UseShellExecute = true;
                Process processChild = Process.Start(info);
                processChild.WaitForExit();
                processChild.Dispose();
            }
            catch(Exception e) {
                Console.WriteLine("PSCreate: " + e.Message);
            }
        }

        private DataTable getSetterMapForANM()
        {
            DataTable dt1=null;
            try
            {
                dt1 = GetDataFromDB("select t.SetterNumber,t.articleid from ldl_trn_articletracker t  where t.publisherid='P1031' and journalid='ANM' and SetterNumber is not null");
            }
            catch(Exception e) {
                Console.WriteLine("getSetterMapForANM: " + e.Message);
            }
            return dt1;
        }

        private void createXMLToHTML(string ArticleID, string dirStartPath)
        {
            try
            {
                string[] folderCode = ArticleID.Split('-');
                if (folderCode[2] == "IPG" || folderCode[2] == "EAG" || folderCode[2] == "TLP" || folderCode[2] == "XPS" || folderCode[2] == "AER" || folderCode[2] == "CTY" || folderCode[2] == "REC")
                {
                    string arg = " -input \"" + dirStartPath + ArticleID + "\\" + folderCode[2] + "-" + folderCode[3] + ".xml\" -output " + "\"" + dirStartPath + ArticleID + "\\" + ArticleID + "_APE.html\"";
                    ProcessStartInfo info = new ProcessStartInfo("D:\\CAPS\\CUP-Online\\CUP-handlers-journal\\CUP_pre_cleanup.exe", arg);
                    info.CreateNoWindow = true;
                    info.UseShellExecute = true;
                    Process processChild = Process.Start(info);
                    processChild.WaitForExit();
                    processChild.Dispose();
                    //processChild1 = Process.Start("D:\\CAPS\\CUP-Online\\CUP-handlers-journal\\CUPJournal_BeforeLoad.exe", " \"" + dirStartPath + ArticleID + "\\" + ArticleID + "_APE.html\"  " + "\"" + dirStartPath + ArticleID + "\\" + ArticleID + "_APE.html\" \"D:\\CAPS\\CUP-Online\\CUP-handlers-journal\\SupportFile\\\"");
                    arg = " \"" + dirStartPath + ArticleID + "\\" + ArticleID + "_APE.html\"  " + "\"" + dirStartPath + ArticleID + "\\" + ArticleID + "_APE.html\" \"D:\\CAPS\\CUP-Online\\CUP-handlers-journal\\SupportFile\\\"";
                    info = new ProcessStartInfo("D:\\CAPS\\CUP-Online\\CUP-handlers-journal\\CUPJournal_BeforeLoad.exe", arg);                    
                    info.CreateNoWindow = true;
                    info.UseShellExecute = true;
                    processChild = Process.Start(info);
                    processChild.WaitForExit();
                    processChild.Dispose();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("createXMLToHTML:: " + e.Message);
            }
        }

        private bool generateWileyMailorDownload(string mailbody, string mailId, string ArticleID, string publisher, string fromMailID,string sub,string GUID,string msgSubject,string mb,string dirStartPath)
        {
            try
            {
                GUID = GUID.Trim();
                string[] jCode = ArticleID.Split('-');
                string journalCode = "";
                string jidC = "";
                
                try
                {
                    if (ArticleID.Contains("_"))
                    {
                        jCode = ArticleID.Split('_');
                        jidC = jCode[0].ToUpper();
                    }
                    else
                    {
                        try
                        {

                            jidC = jCode[2].ToUpper();
                        }
                        catch(Exception e) {
                            if (ArticleID.Contains("-") && jCode.Length == 2)
                            {
                                jidC = jCode[0].ToUpper();
                            }
                            Console.WriteLine(e.Message);
                        }

                    }
                }
                catch { }
                switch (jidC)
                {
                    case "ADEM":
                        journalCode = "aem-journal";
                        break;

                    case "SRIN":
                        journalCode = "srin-journal";
                        break;
                    case "PSSA":
                        journalCode = "pssa-journal";
                        break;

                    case "ENTE":
                        journalCode = "ente";
                        break;

                    case "SOLR":
                        journalCode = "solar-rrl";
                        break;

                    case "PSSB":
                        journalCode = "pssb-journal";
                        break;

                    case "PSSR":
                        journalCode = "pssrrl-journal";
                        break;

                    case "AISY":
                        journalCode = "advintellsyst";
                        break;

                    default:
                        journalCode = "ente";
                        break;
                }

                string taskEndDate = getDataByRegex(mailbody, "Please complete this task by:.*?[0-9]{4}");
                taskEndDate = taskEndDate.Replace("Please complete this task by:", "").Trim();

                mailbody = mailbody.Replace("\r", " ").Replace("\n"," ");
                mailbody = mailbody.Replace("  ", " ");
                // added for wiley
                if (publisher == "P1032" && !fromMailID.ToLower().Contains(".caps") && mailId.ToLower().Contains("wiley.caps@luminad.com"))
                {
                    if (sub.Contains("Task Assignment") && ArticleID.ToUpper().Contains("LDL-") && mailbody.Contains("Dear Lumina Typesetter") && mailbody.ToLower().Contains("rt red"))
                    {
                        string[] folderCode = ArticleID.Split('-');
                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Dear Team,</p><br/><p>Editor has submitted the correction for the article " + ArticleID.Replace("LDL-", "").Replace("WILEY-", "") + " and it has been successfully loaded in CAPS, please deliver the completed revised proof before " + taskEndDate + ".</p><br/><p>Thank You,</p><p>CAPS</p>" + "',getdate()); Update ldl_trn_articletracker set stageid='STM1013' where foldercode='" + ArticleID + "';update ldl_trn_articledetails set WileyRTColor='RED' where publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';";
                        UpdateQuery(dbQuery);
                        if (!Directory.Exists(dirStartPath + ArticleID + "\\Editor\\"))
                        {
                            Directory.CreateDirectory(dirStartPath + ArticleID + "\\Editor\\");
                        }
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");
                        //File.Copy("")

                        updateInXML(dirStartPath + ArticleID + "\\" + folderCode[2].ToUpper() + "-" + folderCode[3] + ".xml", "format_final");
                        versioningOfFile(dirStartPath + ArticleID + "\\RT Red", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", ArticleID);
                        //versioningOfFile(dirStartPath + ArticleID + "\\RT Red", dirStartPath + ArticleID + "\\Editor\\" + journalCode + "_" + GUID + ".go.xml");

                    }
                    else if (sub.Contains("Task Assignment") && ArticleID.ToUpper().Contains("LDL-WILEY-") && mailbody.Contains("Dear Lumina Typesetter") && mailbody.ToLower().Contains("rt yellow"))
                    {
                        // Move to Finalize Pages as suggest by Raja,Manish Tue 8/6/2019 5:41 PM Mail subject: Wiley : 
                        string[] folderCode = ArticleID.Split('-');
                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Dear Team,</p><br/><p>Editor has submitted the correction for the article " + ArticleID.Replace("LDL-", "").Replace("WILEY-", "") + " and it has been successfully loaded in CAPS, please update the correction and deliver the EV Delivery package before " + taskEndDate + ".</p><br/><p>Thank You,</p><p>CAPS</p>" + "',getdate()); Update ldl_trn_articletracker set stageid='STM1016',EditorRound='No' where foldercode='" + ArticleID + "';update ldl_trn_articledetails set WileyRTColor='YELLOW' where publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';";
                        UpdateQuery(dbQuery);

                        if (!Directory.Exists(dirStartPath + ArticleID + "\\Editor\\"))
                        {
                            Directory.CreateDirectory(dirStartPath + ArticleID + "\\Editor\\");
                        }
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");

                        updateInXML(dirStartPath + ArticleID + "\\" + folderCode[2].ToUpper() + "-" + folderCode[3] + ".xml", "Yellow");

                        versioningOfFile(dirStartPath + ArticleID + "\\RT Yellow", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", ArticleID);
                        //versioningOfFile(dirStartPath + ArticleID + "\\RT Yellow", dirStartPath + ArticleID + "\\Editor\\" + journalCode + "_" + GUID + ".go.xml");

                    }
                    else if (sub.Contains("Task Assignment") && ArticleID.ToUpper().Contains("LDL-WILEY-") && mailbody.Contains("Dear Lumina Typesetter") && mailbody.ToLower().Contains("rt green"))
                    {
                        string[] folderCode = ArticleID.Split('-');
                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Dear Team,</p><br/><p>Editor has approved the Online Proof for the article " + ArticleID.Replace("LDL-", "").Replace("WILEY-", "") + " , please deliver the EV Delivery package before " + taskEndDate + ".</p><br/><p>Thank You,</p><p>CAPS</p>" + mb + "',getdate()); Update ldl_trn_articletracker set stageid='STM1016' where foldercode='" + ArticleID + "';update ldl_trn_articledetails set WileyRTColor='GREEN' where publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';";
                        UpdateQuery(dbQuery);
                        try
                        {
                            updateInXML(dirStartPath + ArticleID + "\\" + folderCode[2].ToUpper() + "-" + folderCode[3] + ".xml", "GREEN");

                            downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                            downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");

                            versioningOfFile(dirStartPath + ArticleID + "\\RT Green", dirStartPath + ArticleID + "\\Editor\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", ArticleID);

                            if (Directory.Exists(dirStartPath + ArticleID + "\\EVDelivery\\"))
                            {
                                //Process processChild = Process.Start("D:\\Wiley\\Wiley_EV_Package_1\\Wiley_EV_Package.exe", " \"" + dirStartPath + ArticleID + "\\EVDelivery\\"+"\"");
                            }
                            else
                            {
                                //Process processChild = Process.Start("D:\\Wiley\\Wiley_EV_Package_1\\Wiley_EV_Package.exe", " \"" + dirStartPath + ArticleID+"\\"+"\"");
                            }
                            //processChild.WaitForExit();
                            //processChild.Dispose();

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("generateWileyMailerDownload:" + e.Message);
                        }
                    }
                    else if (sub.ToLower().Contains("rt accepted article") && ArticleID.ToUpper().Contains("LDL-WILEY-") && mailbody.Contains("This task has been cancelled"))
                    {
                        //AA task has been cancelled
                        string[] folderCode = ArticleID.Split('-');
                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1023','STAGE-CHANGE','Mail received from wiley as: AA task has been cancelled','caps1');");
                        UpdateQuery("Update ldl_trn_articletracker set stageid='STM1023' where foldercode='" + ArticleID + "';Update ldl_trn_Articledetails set IsWileyReassigned='Yes' where journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "' and publisherid='P1032'");
                    }
                    else if (sub.ToLower().Contains("rt copyediting article") && ArticleID.ToUpper().Contains("LDL-WILEY-") && mailbody.Contains("This task has been cancelled"))
                    {


                        DataTable dtHoldDetails = GetDataFromDB("select foldercode,stageid,ISONHOLD from ldl_trn_articletracker where foldercode='" + ArticleID + "' ");

                        string foldercode = "";
                        string stage = "";
                        string isOnHold = "";
                        if (dtHoldDetails.Rows.Count == 1)
                        {
                            DataRow dr1 = dtHoldDetails.Rows[0];
                            foldercode = dr1[0].ToString();
                            stage = dr1[1].ToString();
                            isOnHold = dr1[2].ToString();
                            if (isOnHold == "Y" && stage == "STM1003")
                            {
                                // CE task has been cancelled in CE stage
                                string[] folderCode = ArticleID.Split('-');
                                UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1002','STAGE-CHANGE','Mail received from wiley as: CE task has been cancelled in CE stage','caps1');");
                                UpdateQuery("Update ldl_trn_articletracker set stageid='STM1002' where foldercode='" + ArticleID + " and stageid='STM1003'; Update ldl_trn_Articledetails set IsWileyReassigned='Yes' where journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "' and publisherid='P1032'");
                            }
                            else if (stage != "STM1003")
                            {
                                // CE task has been cancelled in non CE stage
                                string[] folderCode = ArticleID.Split('-');
                                UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1023','STAGE-CHANGE','Mail received from wiley as: CE task has been cancelled','caps1');");

                            }
                        }

                    }

                    else if (sub.ToLower().Contains("rt proof correction") && ArticleID.ToUpper().Contains("LDL-WILEY-") && mailbody.Contains("This task has been cancelled"))
                    {
                        string[] folderCode = ArticleID.Split('-');
                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1012','STAGE-CHANGE','Mail received from wiley as: This task has been cancelled');");
                        UpdateQuery("Update ldl_trn_articletracker set stageid='STM1012' where foldercode='" + ArticleID + "';");
                    }



                    else if (sub.Contains("Task Assignment - RT Proof Correction Processing") && ArticleID.ToUpper().Contains("LDL-WILEY-") && GUID != "" && mailbody.Contains("Please download the files"))
                    {
                        string[] folderCode = ArticleID.Split('-');
                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Dear Team,</p><br/><p>Author has submitted the correction for the article " + ArticleID.Replace("LDL-", "").Replace("WILEY-", "") + " and it has been successfully loaded in CAPS, please deliver the completed revised proof before " + taskEndDate + ".</p><br/><p>Thank You,</p><p>CAPS</p>" + mb + "',getdate());  Update ldl_trn_articletracker set stageid='STM1013' where foldercode='" + ArticleID + "';update ldl_trn_articledetails set GUID_author='" + GUID + "' where publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';";
                        UpdateQuery(dbQuery);
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");

                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1013','STAGE-CHANGE');");
                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1012','SUBMITTED')");

                        updateInXML(dirStartPath + ArticleID + "\\" + folderCode[2].ToUpper() + "-" + folderCode[3] + ".xml", "format_final");
                        //format_final
                        versioningOfFile(dirStartPath + ArticleID + "\\RT_Proof_Correction_Processing", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", ArticleID);



                        //versioningOfFile(dirStartPath + ArticleID + "\\RT Proof Correction Processing", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".go.xml");

                    }
                    else if (sub.Contains("Task Assignment - RT Proof Correction Processing") && ArticleID.ToUpper().Contains("LDL-WILEY-") && GUID != "" && mailbody.Contains("Please download the files"))
                    {
                        string[] folderCode = ArticleID.Split('-');
                        if (!Directory.Exists(dirStartPath + ArticleID + "\\AuthorCorrection"))
                        {
                            Directory.CreateDirectory(dirStartPath + ArticleID + "\\AuthorCorrection");
                        }
                        string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Dear Team,</p><br/><p>Author has submitted the correction for the article " + ArticleID.Replace("LDL-", "").Replace("WILEY-", "") + " and it has been successfully loaded in CAPS, please deliver the completed revised proof before " + taskEndDate + ".</p><br/><p>Thank You,</p><p>CAPS</p>" + mb + "',getdate()); Update ldl_trn_articletracker set stageid='STM1013' where foldercode='" + ArticleID + "'; update ldl_trn_articledetails set GUID_author='" + GUID + "' where publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';";
                        UpdateQuery(dbQuery);

                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\AuthorCorrection\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\AuthorCorrection\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");


                        versioningOfFile(dirStartPath + ArticleID + "\\RT Proof Correction Processing", dirStartPath + ArticleID + "\\AuthorCorrection\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", ArticleID);
                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1013','STAGE-CHANGE');");
                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1012','SUBMITTED')");

                    }

                    else if (sub.Contains("Task Assignment - RT Typesetting") && ArticleID.ToUpper().Contains("LDL-WILEY-") && GUID != "" && mailbody.Contains("This manuscript is ready for typesetting. Please typeset it using the available data"))
                    {
                        string[] folderCode = ArticleID.Split('-');
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");

                        string WileyMailDetails = "RT Typesetting";
                        if (!Directory.Exists(dirStartPath + ArticleID + "\\RT Typesetting"))
                        {
                            WileyMailDetails = "RT Typesetting";
                            versioningOfFile(dirStartPath + ArticleID + "\\RT Typesetting", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", ArticleID);
                        }
                        else
                        {
                            versioningOfFile(dirStartPath + ArticleID + "\\TypesettingReassigned", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", ArticleID);
                            WileyMailDetails = "TypesettingReassigned";
                        }
                        string dbquery = "update ldl_trn_articledetails set WileyMailDetails='" + WileyMailDetails + "', GUID_ce='" + GUID + "' where publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';";
                        UpdateQuery(dbquery);

                        //execute if article is on hold and ISONHOLD='Y'
                        DataTable dtHoldDetails = GetDataFromDB("select foldercode,stageid,ISONHOLD from ldl_trn_articletracker where foldercode='" + ArticleID + "' ");

                        if (dtHoldDetails.Rows.Count == 1)
                        {
                            DataRow dr1 = dtHoldDetails.Rows[0];
                            string foldercode = dr1[0].ToString();
                            string stage = dr1[1].ToString();
                            string isOnHold = dr1[2].ToString();
                            if (isOnHold == "Y")
                            {
                                UpdateQuery(" Update ldl_trn_articletracker set ISONHOLD='N' where foldercode='" + ArticleID + "';");
                                UpdateQuery(" Update ldl_trn_ArticleHoldDetails set unholdstatus='UNHOLD',unholduserid='caps1',unholdtime=getdate() where isnull(unholdstatus,'')='' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "'");
                                UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1002','UNHOLD','caps1')");
                            }
                            if (stage == "STM1012")
                            {
                                UpdateQuery("Update ldl_trn_articletracker set stageid='STM1083' where foldercode='" + ArticleID + "';");
                            }
                        }
                    }
                    else if (publisher == "P1032" && !fromMailID.ToLower().Contains(".caps") && mailId.ToLower().Contains("wiley.caps@luminad.com") && !ArticleID.ToUpper().Contains("LDL-WILEY-"))
                    {
                        // check here
                        if (sub.Contains("Task Assignment - RT Copyediting for") && mailbody.Contains("Dear Lumina Typesetter"))
                        {
                            //string[] folderCode = ArticleID.Split('-');
                            string[] copyEditFilesCount = Directory.GetFiles(dirStartPath + "\\NotFound\\WILEY\\UnIdentified\\INBOX\\em\\", "*" + ArticleID.Replace("-", ".") + "*.mail");

                            if (copyEditFilesCount.Length == 0)
                            {
                                string jid = getDataByRegex(mailbody, "The production task GUID is: {.*?}");
                                jid = jid.Replace("The production task GUID is: {", "");
                                jid = jid.Replace("}", "").Trim();
                                GUID = jid.Trim();

                                bool result = uploadZIPToFTP_WileyCE("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", "//Input_folder//Warehouse//" + journalCode + "_" + jid + ".zip", "//Input_folder/" + journalCode + "_" + jid + ".zip", "none");
                                if (result == false)
                                {
                                    return false;
                                }
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                    else if (sub.Contains("Task Assignment - RT Copyediting for") && ArticleID.ToUpper().Contains("LDL-WILEY-") && GUID != "" && mailbody.ToLower().Contains("upload"))
                    {
                        string jid = getDataByRegex(mailbody, "The production task GUID is: {.*?}");
                        jid = jid.Replace("The production task GUID is: {", "");
                        jid = jid.Replace("}", "").Trim();
                        GUID = jid.Trim();

                        string[] folderCode = ArticleID.Split('-');
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");
                        string dbquery = "update ldl_trn_articledetails set GUID_warehouse='" + GUID + "',GUID_ce_reassigned='" + GUID + "' where publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';";
                        UpdateQuery(dbquery);

                        string folder = "RT Copyediting";
                        try
                        {
                            using (ZipArchive archive = ZipFile.OpenRead(dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip"))
                            {
                                foreach (ZipArchiveEntry entry in archive.Entries)
                                {
                                    if (entry.FullName.ToLower().Contains(".tex"))
                                    {
                                        folder = "WILEY_PRETEX";
                                        break;
                                    }
                                }
                            }
                        }
                        catch { }

                        if (!folder.Contains("WILEY_PRETEX"))
                        {

                            // Commented as suggested by Raja, that now onwards we don't need CE reasigned queue now onward July-2019 (Date yaad nahe hai)
                            //Update ldl_trn_articletracker set stageid='STM1076' where foldercode='" + ArticleID + "' and stageid<>'STM1003'

                            // Requrement given by Raja. Update the value if recevied the Pre word in file naming of RT Copyediting task
                            // Mail Sub: CAPS Pending Points - Discussion
                            string wileyPre = "";
                            if (updateForPreInWiley(dirStartPath + ArticleID + "\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip") == true)
                            {
                                wileyPre = "Yes";
                            }
                            // End for Pre mail here...................................................



                            //UpdateQuery("update ldl_trn_articledetails set GUID_warehouse='" + GUID + "' where isnull(GUID_warehouse,'') ='' and publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';");                        
                            // above commented and below added as discussed with RAJA Mon 8/5/2019 7:02 PM mail subject : FW: pssa.201900406 - Not closed in EM
                            UpdateQuery("update ldl_trn_articledetails set GUID_warehouse='" + GUID + "',IsWileyPreArticle='" + wileyPre + "' where  publisherid='P1032' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "';");

                            string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','vch.production@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>Dear Team,</p><br/><p>Editor has assigned article " + ArticleID.Replace("LDL-", "").Replace("WILEY-", "") + " and it has been successfully loaded in CAPS, please deliver it before " + taskEndDate + ".</p><br/><p>Thank You,</p><p>CAPS</p>" + mb + "',getdate());";
                            UpdateQuery(dbQuery);

                            versioningOfFile(dirStartPath + ArticleID + "\\RT Copyediting", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", ArticleID);
                            //versioningOfFile(dirStartPath + ArticleID + "\\RT Copyediting", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".go.xml");


                            string[] ard = ArticleID.Split('-');
                            //0- LDL 1- PID 2- JID 3- AID
                            if (ard.Length == 4)
                            {
                                try
                                {
                                    bool result = extractFiles(dirStartPath + ArticleID + "\\", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", ArticleID);
                                    string bundlePath = dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip";
                                    if (result == true)
                                    {
                                        bundlePath = dirStartPath + ArticleID + "\\temp\\" + journalCode + "_" + GUID + ".zip";
                                    }

                                    versioningofDownloadFiles(dirStartPath + ArticleID + "\\", "JOB ANALYSIS", ard[2] + "_" + ard[3], "CEBundle", bundlePath, ard[2] + "_" + ard[3] + "_JA_Input.zip", true);
                                    if (result == true)
                                    {
                                        removeTempDirAndFiles(dirStartPath + ArticleID + "\\temp");
                                    }

                                    if (File.Exists(bundlePath))
                                    {
                                        FileInfo f = new FileInfo(bundlePath);
                                        removeTempDirAndFiles(dirStartPath + ArticleID + "\\" + f.Name.Replace(".zip", ""));
                                    }
                                    if (File.Exists(dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID))
                                    {
                                        removeTempDirAndFiles(dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID);
                                    }

                                    DataTable dtHoldDetails = GetDataFromDB("select foldercode,stageid,ISONHOLD from ldl_trn_articletracker where foldercode='" + ArticleID + "' ");

                                    string foldercode = "";
                                    string stage = "";
                                    string isOnHold = "";
                                    if (dtHoldDetails.Rows.Count == 1)
                                    {
                                        DataRow dr1 = dtHoldDetails.Rows[0];
                                        foldercode = dr1[0].ToString();
                                        stage = dr1[1].ToString();
                                        isOnHold = dr1[2].ToString();
                                        if (isOnHold == "Y")
                                        {
                                            UpdateQuery("Update ldl_trn_articletracker set isonhold='N' where foldercode='" + ArticleID + "';");
                                            UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'" + stage + "','STAGE-CHANGE','Hold removed, mail received from wiley as: CE integeration','caps1');");
                                            UpdateQuery("Update ldl_trn_ArticleHoldDetails set unholdtime=getdate(),unholdstatus='UNHOLD',unholduserid='caps1',unholdcomment='Hold removed, mail received from wiley as: CE integeration' where journalid='" + ard[2] + "' and articleid='" + ard[3] + "' and publisherid='P1032' and unholdstatus is null;");
                                        }
                                    }
                                    UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'" + stage + "','','Mail received from wiley as: CE integeration','caps1');");

                                    
                                }
                                catch { }
                            }


                        }
                        else
                        {
                            // For PreTex after versioning: send mail to PreTeX team and moved artilce in PreTeX stage
                            versioningOfFile(dirStartPath + ArticleID + "\\" + folder + "", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", ArticleID);
                            updateForPreTeX(ArticleID);
                        }


                        try
                        {
                            if (File.Exists(dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip"))
                            {
                                if (!Directory.Exists(dirStartPath + ArticleID + "\\source"))
                                {
                                    Directory.CreateDirectory(dirStartPath + ArticleID + "\\source");
                                }
                                if (!Directory.Exists(dirStartPath + ArticleID + "\\source\\CE"))
                                {
                                    Directory.CreateDirectory(dirStartPath + ArticleID + "\\source\\CE");
                                }
                                File.Copy(dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", dirStartPath + ArticleID + "\\source\\CE\\" + journalCode + "_" + GUID + ".zip", true);
                            }
                        }
                        catch { }

                        DataTable dtHoldDtls = GetDataFromDB("select foldercode,stageid,ISONHOLD from ldl_trn_articletracker where foldercode='" + ArticleID + "' ");

                        if (dtHoldDtls.Rows.Count == 1)
                        {
                            DataRow dr1 = dtHoldDtls.Rows[0];
                            string foldercode = dr1[0].ToString();
                            string stage = dr1[1].ToString();
                            string isOnHold = dr1[2].ToString();
                            if (isOnHold == "Y")
                            {
                                UpdateQuery(" Update ldl_trn_articletracker set ISONHOLD='N' where foldercode='" + ArticleID + "';");
                                UpdateQuery(" Update ldl_trn_ArticleHoldDetails set unholdstatus='UNHOLD',unholduserid='caps1',unholdtime=getdate() where isnull(unholdstatus,'')='' and journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "'");
                                UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1002','UNHOLD','caps1')");
                            }
                        }

                    }

                    else if (sub.Contains("Task Assignment - Additional typesetting request") && ArticleID.ToUpper().Contains("LDL-WILEY-") && GUID != "" && mailbody.Contains("The task Additional typesetting request has been assigned"))
                    {
                        string[] ard = ArticleID.Split('-');
                        if (!Directory.Exists(dirStartPath + ArticleID + "\\AdditionalTypesetting"))
                        {
                            Directory.CreateDirectory(dirStartPath + ArticleID + "\\AdditionalTypesetting");
                        }
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\AdditionalTypesetting\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none", "");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\AdditionalTypesetting\\" + journalCode.Replace("solar-rrl", "solar-rrl") + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none", "");

                        UpdateQuery("Update ldl_trn_articletracker set stageid='STM1016' where foldercode='" + ArticleID + "';");
                        UpdateQuery("Update ldl_trn_articledetials set isAdditionalTypesetting='Yes' where articleid='"+ard[3]+"' and journalid='"+ard[2]+"'");
                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'STM1016','Moved to Finalize pages as additional typesetting mail received','caps1')");
                    }
                    else if (sub.Contains("Task Assignment - Additional typesetting request") && !ArticleID.ToUpper().Contains("LDL-WILEY-") && GUID != "" && mailbody.Contains("The task Additional typesetting request has been assigned"))
                    {
                        string jid = getDataByRegex(mailbody, "The production task GUID is: {.*?}");
                        jid = jid.Replace("The production task GUID is: {", "");
                        jid = jid.Replace("}", "").Trim();
                        GUID = jid.Trim();

                        bool result = uploadZIPToFTP_WileyCE("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", "//Input_folder//Warehouse//" + journalCode + "_" + jid + ".zip", "//Input_folder/" + journalCode + "_" + jid + ".zip", "none");
                        if (result == false)
                        {
                            return false;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return true;
        }

        private void updateForPreTeX(string ArticleID)
        {
            try
            {
                string[] articleDetails = ArticleID.Split('-');
                UpdateQuery("Update ldl_trn_articletracker set stageid='STM1088',INPUTTYPE='TeX' where foldercode='" + ArticleID + "'");
                UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark) values ('P1032','" + articleDetails[2] + "','" + articleDetails[3] + "',getdate(),'STM1088','STAGE-CHANGE','TeX article, as TeX exist in package.');");
                //UpdateQuery("insert into ldl_trn_ArticleMovementRegister_CUP ");

                StringBuilder builder = new StringBuilder();
                builder.Append("Dear Team,");
                builder.Append("<br>");
                builder.Append("<br>");
                builder.Append("<p>" + articleDetails[2] + "-" + articleDetails[3] + " is ready for PreTeX process, kindly complete this and move to next stage within 4 Hours.</p>");
                builder.Append("Thank you");
                builder.Append("<br>");
                builder.Append("Regards,");
                builder.Append("<br>");
                builder.Append("Lumina CAPS Technical Team");

                //select tomailid,ccmailid from ldl_mst_MailIDForWarehouse w where w.publisherid='WILEYPreTeX'
                DataTable mailDetailsForTeX = GetDataFromDB("select tomailid,ccmailid from ldl_mst_MailIDForWarehouse w where w.publisherid='WILEYPreTeX'");

                if (mailDetailsForTeX.Rows.Count == 1)
                {
                    DataRow dr1 = mailDetailsForTeX.Rows[0];
                    string toMailID = dr1[0].ToString();
                    string ccMailID = dr1[1].ToString();

                    string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','" + toMailID + "','" + ccMailID + "','" + "Wiley-VCH: " + articleDetails[3] + "-" + articleDetails[2] + " is ready for PreTeX process" + "','" + builder.ToString() + "',getdate())";
                    UpdateQuery(dbQuery);
                }
            }
            catch { }
        }

        public bool extractFiles(string path, string zipPath, string ArticleID)
        {
            try
            {                
                string[] articleDetails=ArticleID.Split('-');
                FileInfo f = new FileInfo(zipPath);
                if (!Directory.Exists(path + f.Name.Replace(".zip", "")))
                {
                    Directory.CreateDirectory((path + f.Name.Replace(".zip", "")));
                }

                if (Directory.Exists(path + "\\temp\\"))
                {
                    removeTempDirAndFiles(path + "\\temp\\");
                }
                if(!Directory.Exists(path+"\\temp")){
                    Directory.CreateDirectory(path+"\\temp\\");
                }

                if(Directory.Exists(path + f.Name.Replace(".zip", "")))
                {
                    removeTempDirAndFiles(path + f.Name.Replace(".zip", ""));
                }

                string extractPath = path + f.Name.Replace(".zip", "");
                ZipFile.ExtractToDirectory(zipPath, extractPath+"\\");
                
                string[] files = Directory.GetFiles(extractPath+"\\", "*.doc*");

                string[] filesTeX = Directory.GetFiles(extractPath + "\\", "*.tex");
                if (filesTeX.Length > 0)
                {
                    UpdateQuery("Update ldl_trn_articletracker set stageid='STM1088' where foldercode='" + ArticleID + "'");
                    UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark) values ('P1032','" + articleDetails[2] + "','" + articleDetails[3] + "',getdate(),'STM1088','STAGE-CHANGE','TeX article, as TeX exist in package.');");
                    //UpdateQuery("insert into ldl_trn_ArticleMovementRegister_CUP ");

                    StringBuilder builder = new StringBuilder();
                    builder.Append("Dear Sir,");
                    builder.Append("<br>");
                    builder.Append("<br>");
                    builder.Append("<p>" + articleDetails[2] + "-" + articleDetails[3] + " is ready for PreTeX process, kindly complete this and move to next stage within 4 Hours.</p>");
                    builder.Append("Thank you");
                    builder.Append("<br>");
                    builder.Append("Regards,");
                    builder.Append("<br>");
                    builder.Append("Lumina CAPS Technical Team");

                    //select tomailid,ccmailid from ldl_mst_MailIDForWarehouse w where w.publisherid='WILEYPreTeX'
                    DataTable mailDetailsForTeX = GetDataFromDB("select tomailid,ccmailid from ldl_mst_MailIDForWarehouse w where w.publisherid='WILEYPreTeX'");

                    if (mailDetailsForTeX.Rows.Count == 1)
                     {
                         DataRow dr1 = mailDetailsForTeX.Rows[0];
                         string toMailID = dr1[0].ToString();
                         string ccMailID = dr1[1].ToString();

                         string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1032','wiley.caps@luminad.com','" + toMailID + "','" + ccMailID + "','" + "Wiley-VCH: " + articleDetails[3] + "-" + articleDetails[2] + " is ready for PreTeX process" + "','"+builder.ToString()+"',getdate())";
                         UpdateQuery(dbQuery);
                     }

                }
                
                long largestSize = 0;
                string largeFile = "";
                for (int i = 0; i < files.Length; i++)
                {
                    FileInfo fl = new FileInfo(files[i]);
                    if (fl.Length > largestSize)
                    {
                        largestSize = fl.Length;
                        largeFile = files[i];
                    }
                }


               



                if (!Directory.Exists(path + "AcceptedArticle"))
                {
                    Directory.CreateDirectory(path + "AcceptedArticle");
                }
                FileInfo fileToCopy = new FileInfo(largeFile);
                
                String fName = fileToCopy.Name;

                fileToCopy.CopyTo(largeFile.Replace(fName, articleDetails[2] + "-" + articleDetails[3] + "_msp" + fileToCopy.Extension),true);
                fileToCopy.MoveTo(largeFile.Replace(fName,articleDetails[2]+"-"+articleDetails[3]+fileToCopy.Extension));

                //string[] dirs = Directory.GetDirectories((path + "AcceptedArticle"), "*Version*");
                //if (!Directory.Exists(path + "AcceptedArticle\\Version" + Convert.ToString(dirs.Length+1)+"\\Input"))
                //{
                //    Directory.CreateDirectory(path + "AcceptedArticle\\Version" + Convert.ToString(dirs.Length+1) + "\\Input");
                //}
                //string newPath = path + "AcceptedArticle\\Version" + Convert.ToString(dirs.Length + 1) + "\\Input\\";

                // comment below
                //fileToCopy.CopyTo(path + "AcceptedArticle\\" + "\\manuscript" + fileToCopy.Extension, true);                
                //fileToCopy.CopyTo(newPath + "\\" + articleDetails[2] + "-" + articleDetails[3] + "_msp"  + fileToCopy.Extension, true);                

                //JMSProducer jProd = new JMSProducer();
                //jProd.postMessage("Warehouse" + ArticleID, "Redline");

                ZipFile.CreateFromDirectory(extractPath, path + "\\temp\\" + f.Name);
                if (File.Exists(path + "\\temp\\" + f.Name))
                    return true;

                //Console.ReadLine();
            }
            catch(Exception e) {
                Console.WriteLine(e.Message);               
            }
            return false;
        }

        public void removeTempDirAndFiles(string folderPath)
        {
            try
            {
                if(Directory.Exists(folderPath)){
                    string[] fileToDelete = Directory.GetFiles(folderPath);
                for (int i = 0; i < fileToDelete.Length; i++)
                {
                    FileInfo fin = new FileInfo(fileToDelete[i]);
                    fin.Delete();
                }
                Directory.Delete(folderPath);
                }
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private bool updateForPreInWiley(string fileName)
        {
            try
            {
                if (File.Exists(fileName))
                {
                    using (ZipArchive archive = ZipFile.OpenRead(fileName))
                    {
                        //archive.Entries[1].ToString()
                        for (int i = 0; i < archive.Entries.Count; i++)
                        {
                            if (archive.Entries[i].ToString().ToLower().Contains("_pre"))
                            {
                                return true;
                            }
                        }
                    }
                }                
            }
            catch { }
            return false;

        }


        private string downloadWileyAMPackage(string mailbody, string mailId, string ArticleID, string publisher, string fromMailID, string sub, string GUID, string msgSubject, string mb, string dirStartPath)
        {
            try
            {
                 if (!dirStartPath.EndsWith("\\"))
                {
                    dirStartPath= dirStartPath + "\\";
                }

                GUID = GUID.Trim();
                string[] jCode = ArticleID.Split('-');
                string journalCode = "";
                string jidC = "";

                try
                {
                    if (ArticleID.Contains("_"))
                    {
                        jCode = ArticleID.Split('_');
                        jidC = jCode[0].ToUpper();
                    }
                    else
                    {
                        try
                        {

                            jidC = jCode[2].ToUpper();
                        }
                        catch (Exception e)
                        {
                            if (ArticleID.Contains("-") && jCode.Length == 2)
                            {
                                jidC = jCode[0].ToUpper();
                            }
                            Console.WriteLine(e.Message);
                        }

                    }
                }
                catch { }
                switch (jidC)
                {
                    case "ADEM":
                        journalCode = "aem-journal";
                        break;

                    case "SRIN":
                        journalCode = "srin-journal";
                        break;
                    case "PSSA":
                        journalCode = "pssa-journal";
                        break;

                    case "ENTE":
                        journalCode = "ente";
                        break;

                    case "SOLR":
                        journalCode = "solar-rrl";
                        break;

                    case "PSSB":
                        journalCode = "pssb-journal";
                        break;

                    case "PSSR":
                        journalCode = "pssrrl-journal";
                        break;

                    case "AISY":
                        journalCode = "advintellsyst";
                        break;

                    default:
                        journalCode = "ente";
                        break;
                }

                string taskEndDate = getDataByRegex(mailbody, "Please complete this task by:.*?[0-9]{4}");
                taskEndDate = taskEndDate.Replace("Please complete this task by:", "").Trim();

                mailbody = mailbody.Replace("\r", " ").Replace("\n", " ");
                mailbody = mailbody.Replace("  ", " ");
                // added for wiley
                if (publisher == "P1032" && !fromMailID.ToLower().Contains(".caps") && mailId.ToLower().Contains("wiley.caps@luminad.com"))
                {
                    if (sub.Contains("Task Assignment") && ArticleID.ToUpper().Contains("LDL-") && mailbody.Contains("Dear Lumina Typesetter") && mailbody.ToLower().Contains("rt accepted article"))
                    {
                        string[] folderCode = ArticleID.Split('-');
                        if (!Directory.Exists(dirStartPath + ArticleID + "\\AM\\"))
                        {
                            Directory.CreateDirectory(dirStartPath + ArticleID + "\\AM\\");
                        }
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", "//Input_folder/" + journalCode + "_" + GUID + ".zip", "none","");
                        downloadZip("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".go.xml", "//Input_folder/" + journalCode + "_" + GUID + ".go.xml", "none","");

                        //File.Copy("")

                        versioningOfFile(dirStartPath + ArticleID + "\\AM", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", ArticleID);

                       

                        string[] ard=ArticleID.Split('-');
                        //0- LDL 1- PID 2- JID 3- AID
                        if (ard.Length == 4)
                        {
                            versioningofDownloadFiles(dirStartPath + ArticleID + "\\", "JOB ANALYSIS", ard[2] + "_" + ard[3], "AABundle", dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", ard[2] + "_" + ard[3] + "_JA_Input.zip", true);
                        }

                        string foldercode = "";
                        string stage = "";
                        string isOnHold = "";
                        DataTable dtHoldDtls = GetDataFromDB("select foldercode,stageid,ISONHOLD from ldl_trn_articletracker where foldercode='" + ArticleID + "' ");
                        if (dtHoldDtls.Rows.Count == 1)
                        {
                            DataRow dr1 = dtHoldDtls.Rows[0];
                            foldercode = dr1[0].ToString();
                            stage = dr1[1].ToString();
                            isOnHold = dr1[2].ToString();
                            if (isOnHold == "Y")
                            {
                                UpdateQuery("Update ldl_trn_articletracker set isonhold='N' where foldercode='" + ArticleID + "';");
                                UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'STM1023','STAGE-CHANGE','Hold removed, mail received from wiley as: AA integeration','caps1');");
                                UpdateQuery("Update ldl_trn_ArticleHoldDetails set unholdtime=getdate(), unholdstatus='UNHOLD',unholduserid='caps1',unholdcomment='Hold removed, mail received from wiley as: AA integeration' where journalid='" + ard[2] + "' and articleid='" + ard[3] + "' and publisherid='P1032' and unholdstatus is null;");
                            }
                        }
                        UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'" + stage + "','','Mail received from wiley as: AA integeration','caps1');");

                        try
                        {
                            if (!Directory.Exists(dirStartPath + ArticleID + "\\AA"))
                            {
                                Directory.CreateDirectory(dirStartPath + ArticleID + "\\AA");
                            }
                            if (!Directory.Exists(dirStartPath + ArticleID + "\\source\\AA"))
                            {
                                Directory.CreateDirectory(dirStartPath + ArticleID + "\\source\\AA");
                            }
                            File.Copy(dirStartPath + ArticleID + "\\" + journalCode + "_" + GUID + ".zip", dirStartPath + ArticleID + "\\source\\AA\\" + journalCode + "_" + GUID + ".zip",true);
                        }
                        catch { }
                        if(File.Exists(dirStartPath + ArticleID +"\\" +  journalCode + "_" + GUID + ".zip")){
                            return "success";
                        }
                    }
                    else if (sub.Contains("Task Assignment") && !ArticleID.ToUpper().Contains("LDL-") && mailbody.Contains("Dear Lumina Typesetter") && mailbody.ToLower().Contains("rt accepted article"))
                    {
                        string jid = getDataByRegex(mailbody, "The production task GUID is: {.*?}");
                        jid = jid.Replace("The production task GUID is: {", "");
                        jid = jid.Replace("}", "").Trim();
                        GUID = jid.Trim();
                        bool result = uploadZIPToFTP_WileyCE("ftp.luminadus.com", "21", "Wiley-caps", "Lumina@755", "//Input_folder//Warehouse//" + journalCode + "_" + jid + ".zip", "//Input_folder/" + journalCode + "_" + jid + ".zip", "none");
                        if(result==false){
                            return "fail";
                        }
                        
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return "success";
        }

        public void versioningofDownloadFiles(string path, string rootJobPath, string rootZipPath, string folderCode,string filePath,string zipFileName,bool zipReq)
        {
            
            try
            {                
                string dir = "";
                string bundleName = "";
                if (folderCode == "AABundle")
                {
                    bundleName = "CEBundle";
                }
                else if (folderCode == "CEBundle")
                {
                    bundleName = "AABundle";
                }
                
                string folderName = path + rootJobPath;
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }
                string[] dirs = Directory.GetDirectories(folderName, "*Version*");
                dir = folderName + "\\Version1\\Input\\" + rootZipPath + "\\" + folderCode + "\\";
                string dir2 = folderName + "\\Version1\\Input\\" + rootZipPath + "\\" + bundleName + "\\";

                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName + "\\Version1\\Input\\" + rootZipPath + "\\" + folderCode + "\\");
                }
                else
                {
                    if (Directory.Exists(dir))
                    {
                        string[] files = Directory.GetFiles(dir);
                        if (files.Length > 0)
                        {
                            if (!Directory.Exists(folderName + "\\Version" + Convert.ToString(dirs.Length) + "\\Input\\" + rootZipPath + "\\" + folderCode + "\\"))
                            {
                                dir = folderName + "\\Version" + Convert.ToString(dirs.Length) + "\\Input\\" + rootZipPath + "\\" + folderCode + "\\";
                                dir2 = folderName + "\\Version" + Convert.ToString(dirs.Length) + "\\Input\\" + rootZipPath + "\\" + bundleName + "\\";
                                Directory.CreateDirectory(folderName + "\\Version" + Convert.ToString(dirs.Length) + "\\Input\\" + rootZipPath + "\\" + folderCode + "\\");
                            }
                            else
                            {
                                dir = folderName + "\\Version" + Convert.ToString(dirs.Length + 1) + "\\Input\\" + rootZipPath + "\\" + folderCode + "\\";
                                dir2 = folderName + "\\Version" + Convert.ToString(dirs.Length + 1) + "\\Input\\" + rootZipPath + "\\" + bundleName + "\\";
                                Directory.CreateDirectory(folderName + "\\Version" + Convert.ToString(dirs.Length + 1) + "\\Input\\" + rootZipPath + "\\" + folderCode + "\\");
                            }
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(dir);
                    }                    
                }

                // Copying file to Version after checking if directory exist
                if (File.Exists(filePath) && Directory.Exists(dir))
                {
                    FileInfo f = new FileInfo(filePath);
                    f.CopyTo(dir + f.Name.ToString(), true);
                    //f.CopyTo(dir + f.Name.ToString().Replace(".zip", ".go.xml"), true);
                }

                //1. update the zip path remaiaing
                //2. AABundle CEBundle

                string[] aaFiles = Directory.GetFiles(dir);
                if (zipReq == true)
                {
                   
                    if (Directory.Exists(dir.Replace(folderCode+"\\",bundleName+"\\")))
                    {
                        string[] ceFiles = Directory.GetFiles(dir.Replace(folderCode, bundleName));
                        if (aaFiles.Length > 0 && ceFiles.Length > 0)
                        {
                            string zipPath = "";
                            string startPath = dir.Replace("\\AABundle", "").Replace("\\CEBundle", "");                            
                            string parentDir = Directory.GetParent(dir.Replace("\\AABundle", "").Replace("\\CEBundle","")).FullName;
                            parentDir = Directory.GetParent(parentDir).FullName;
                            zipPath = parentDir + "\\" + zipFileName;
                            if (File.Exists(zipPath)) {
                                File.Delete(zipPath);
                            }
                            ZipFile.CreateFromDirectory(startPath, zipPath);
                        }
                    }
                }
            }
            catch(Exception e) {
                Console.WriteLine(e.Message);
            }
            
        }

        public void versioningOfFile(string folderName, string filePath, string ArticleID)
        {
            try
            {
                string dir = "";
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName + "\\Version1\\Input");
                    if (folderName.Contains("WILEY_PRETEX") || folderName.Contains("TypesettingReassigned") || folderName.Contains("RT_Proof_Correction_Processing") || folderName.Contains("Red") || folderName.Contains("Yellow") || folderName.Contains("Green"))
                    {
                        Directory.CreateDirectory(folderName + "\\Version1\\Output");
                    }
                    dir = folderName + "\\Version1\\Input\\";
                }
                else
                {
                    string[] dirs = Directory.GetDirectories(folderName);
                    if (dirs.Length == 0)
                    {
                        Directory.CreateDirectory(folderName + "\\Version1\\Input");
                        if (folderName.Contains("WILEY_PRETEX") || folderName.Contains("TypesettingReassigned") || folderName.Contains("RT_Proof_Correction_Processing") || folderName.Contains("Red") || folderName.Contains("Yellow") || folderName.Contains("Green"))
                        {
                            Directory.CreateDirectory(folderName + "\\Version1\\Output");
                        }
                        dir = folderName + "\\Version1\\Input\\";
                    }
                    else
                    {

                        Directory.CreateDirectory(folderName + "\\Version"+ Convert.ToString(dirs.Length+1)+"\\Input");
                        if (folderName.Contains("WILEY_PRETEX") || folderName.Contains("TypesettingReassigned") || folderName.Contains("RT_Proof_Correction_Processing") || folderName.Contains("Red") || folderName.Contains("Yellow") || folderName.Contains("Green"))
                        {
                            //Directory.CreateDirectory(folderName + "\\Version" + Convert.ToString(dirs.Length + 1) + "\\Input");
                            Directory.CreateDirectory(folderName + "\\Version" + Convert.ToString(dirs.Length + 1) + "\\Output");

                        }

                        dir = folderName + "\\Version" + Convert.ToString(dirs.Length + 1) + "\\Input\\";
                    }
                }

                if (File.Exists(filePath) && Directory.Exists(dir))
                {
                    FileInfo f = new FileInfo(filePath);
                    try
                    {
                        //If not RT_GREEN then copy zip and go.xml
                        if (!folderName.ToLower().Contains("green")){
                            if (folderName.Contains("WILEY_PRETEX"))
                            {
                                string[] ard = ArticleID.Split('-');
                                f.CopyTo(dir + ard[2] + "-" +ard[3]+ "_PRETEX.zip", true);

                                if(File.Exists(filePath.Replace(".zip", ".go.xml")))
                                {
                                    FileInfo f1 = new FileInfo(filePath.Replace(".zip", ".go.xml"));
                                    f1.CopyTo(dir + f1.Name.ToString(), true);

                                }
                            }
                            else
                            {
                                f.CopyTo(dir + f.Name.ToString(), true);
                                //f.CopyTo(dir + f.Name.ToString().Replace(".zip", ".go.xml"), true);
                                if (File.Exists(filePath.Replace(".zip", ".go.xml")))
                                {
                                    FileInfo f1 = new FileInfo(filePath.Replace(".zip", ".go.xml"));
                                    f1.CopyTo(dir + f1.Name.ToString(), true);

                                }
                            }
                            
                        }                        
                    }
                    catch(Exception e) {
                        Console.WriteLine(e.Message);
                    }
                    try
                    {

                        moveFilesToVersioning(f, dir, folderName, ArticleID);

                    }
                    catch { }                    
                }
            }
            catch { }
        }

        private void moveFilesToVersioning(FileInfo f, string dir, string folderName, string ArticleID)
        {
            try
            {
                DirectoryInfo dir11 = new DirectoryInfo(f.Directory.ToString());
                
                string dir1 = dir11.Parent.FullName;
                if(dir1.EndsWith("stock")){
                    dir1 = f.DirectoryName;
                }

                if (folderName.Contains("TypesettingReassigned"))
                {
                    //FileInfo f1 = new FileInfo(filePath);
                    //DirectoryInfo dir1 = new DirectoryInfo(f.Directory.ToString());
                    //string dirName = dir1.Name;
                    string[] AID = ArticleID.Split('-');
                    if (File.Exists(dir1+ "\\" + AID[2].ToLower() + AID[3] + ".xml"))
                    {
                        File.Copy(dir1 + "\\" + AID[2].ToLower() + AID[3] + ".xml", dir + AID[2].ToLower() + AID[3] + ".xml", true);
                    }
                }
                else if (folderName.ToUpper().Contains("RED"))
                {
                    // FileInfo f1 = new FileInfo(filePath);
                    //DirectoryInfo dir1 = new DirectoryInfo(f.Directory.ToString());
                    //string dirName = dir1.Name;
                    string[] AID = ArticleID.Split('-');


                    if (File.Exists(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml"))
                    {
                         string[] zipFiles = Directory.GetFiles(dir,"*.zip");

                         if (zipFiles.Length == 1)
                         {
                             zippingVersionFile(dir, AID, dir1, zipFiles);

                         }
                         else
                         {

                             File.Copy(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml", dir + AID[2] + "-" + AID[3] + ".xml", true);
                         }
                    }
                }
                else if (folderName.ToUpper().Contains("YELLOW"))
                {
                    //FileInfo f1 = new FileInfo(filePath);
                    //DirectoryInfo dir1 = new DirectoryInfo(f.Directory.ToString());
                    //string dirName = dir1.Name;
                    string[] AID = ArticleID.Split('-');
                    if (File.Exists(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml"))
                    {
                        File.Copy(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml", dir + AID[2] + "-" + AID[3] + ".xml", true);
                    }
                }
                else if (folderName.ToUpper().Contains("GREEN"))
                {
                    //FileInfo f1 = new FileInfo(filePath);
                    //DirectoryInfo dir1 = new DirectoryInfo(f.Directory.ToString());
                    //string dirName = dir1.Name;
                    string[] AID = ArticleID.Split('-');
                    if (File.Exists(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml"))
                    {
                        File.Copy(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml", dir + AID[2] + "-" + AID[3] + ".xml", true);
                    }
                }
                else if (folderName.Contains("RT_Proof_Correction_Processing"))
                {
                    //FileInfo f1 = new FileInfo(filePath);
                    //DirectoryInfo dir1 = new DirectoryInfo(f.Directory.ToString());
                    //string dirName = dir1.Name;
                    string[] AID = ArticleID.Split('-');
                    if (File.Exists(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml"))
                    {
                        

                        //if (File.Exists(dir1 + "\\" + AID[2] + "-" + AID[3] + ".pdf"))
                        //{
                        //    File.Copy(dir1 + "\\" + AID[2] + "-" + AID[3] + ".pdf", dir + AID[2] + "-" + AID[3] + ".pdf", true);
                        //}

                       
                        // Zipping
                        string[] zipFiles = Directory.GetFiles(dir,"*.zip");

                        if (zipFiles.Length == 1)
                        {
                            zippingVersionFile(dir, AID, dir1, zipFiles);

                        }
                        else
                        {
                            File.Copy(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml", dir + AID[2] + "-" + AID[3] + ".xml", true);
                        }                        


                        //string txt = File.ReadAllText(dir + AID[2] + "-" + AID[3] + ".xml", Encoding.UTF8);
                        //txt = txt.Replace("pmg:currentqueue=\"format_proof\"", "pmg:currentqueue=\"format_final\" ");
                        //File.WriteAllText(dir + AID[2] + "-" + AID[3] + ".xml", txt, Encoding.UTF8);
                    }
                    else if (File.Exists(dir1 + "\\" + AID[2].ToLower() + AID[3] + ".xml"))
                    {
                         string[] zipFiles = Directory.GetFiles(dir,"*.zip");

                         if (zipFiles.Length == 1)
                         {
                             zippingVersionFile(dir, AID, dir1, zipFiles);

                         }
                         else
                         {
                             File.Copy(dir1 + "\\" + AID[2].ToLower() + AID[3] + ".xml", dir + AID[2].ToLower() + AID[3] + ".xml", true);
                         }
                        //string txt = File.ReadAllText(dir + AID[2].ToLower() + AID[3] + ".xml", Encoding.UTF8);
                        //txt = txt.Replace("pmg:currentqueue=\"format_proof\"", "pmg:currentqueue=\"format_final\" ");
                        //File.WriteAllText(dir + AID[2].ToLower() + AID[3] + ".xml", txt, Encoding.UTF8);
                    }

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void zippingVersionFile(string dir, string[] AID, string dir1, string[] zipFiles)
        {
            try
            {
                string extractPath = dir + AID[2] + "-" + AID[3] + "_corr";
                ZipFile.ExtractToDirectory(zipFiles[0], extractPath);
                if (Directory.Exists(extractPath))
                {
                    string[] xmlFile = Directory.GetFiles(extractPath, "*.xml");

                    for (int k = 0; k < xmlFile.Length; k++)
                    {
                        try
                        {
                            File.Delete(xmlFile[k]);
                        }
                        catch { }
                    }

                    File.Copy(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml", extractPath + "\\" + AID[2] + "-" + AID[3] + ".xml", true);
                    ZipFile.CreateFromDirectory(extractPath, dir + AID[2] + "-" + AID[3] + "_corr.zip");

                    removeTempDirAndFiles(extractPath);
                }
                else
                {
                    File.Copy(dir1 + "\\" + AID[2] + "-" + AID[3] + ".xml", dir + AID[2] + "-" + AID[3] + ".xml", true);
                }
            }
            catch { }
        }

        private void updateInXML(string fileName,string color)
        {
            try
            {
                if (File.Exists(fileName))
                {
                    StreamReader reader = new StreamReader(fileName);
                    string input = reader.ReadToEnd();
                    reader.Close();

                    string currentTime = string.Format("{0:yyyy-MM-dd-hhmmss}", DateTime.Now);
                    string bkFileName = fileName.Replace(".xml", "_bkup" + currentTime + ".xml");
                    File.Copy(fileName, bkFileName);

                    //if (File.Exists(fileName) && File.Exists(fileName.Replace(".xml", "_bkup.xml")))
                    //{
                    //    File.Delete(fileName);
                    //}
                    if(File.Exists(bkFileName)){
                        File.WriteAllText(fileName, String.Empty);

                        using (StreamWriter writer = new StreamWriter(fileName, true))
                        {
                            {
                                string data = getDataByRegex(input, "pmg:currentqueue=\".*?\"");
                                string output = "";
                                if (color.ToLower() == "green" || color.ToLower() == "yellow")
                                {
                                    output = input.Replace(data, "pmg:currentqueue=\"publish_online\"");
                                }
                                else if (color.ToLower() == "format_final")
                                {
                                    output = input.Replace(data, "pmg:currentqueue=\"format_final\"");
                                }
                                else
                                {
                                    output = input.Replace(data, "pmg:currentqueue=\"issue_article\"");
                                }

                                writer.Write(output);
                            }
                            writer.Close();
                        }
                    }
                    
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public Boolean updationForCUPFirstProof(string publisher, string ArticleID)
        {

            Boolean result = false;
            try
            {
                DataTable dt1 = GetDataFromDB("Select publisherid,journalid,articleid,stageid from ldl_trn_articletracker where FOLDERCODE='" + ArticleID + "' and stageid='STM1012'");
                UpdateQuery("Update ldl_trn_Articletracker set stageid='STM1013' where foldercode='" + ArticleID + "' and stageid='STM1012'");
                if (dt1.Rows.Count == 1)
                {
                    DataRow dr1 = dt1.Rows[0];
                    string publisherid = dr1[0].ToString();
                    string journalID = dr1[1].ToString();
                    string articleid = dr1[2].ToString();
                    string articleStageID = dr1[3].ToString();

                    if (articleStageID == "STM1012")
                    {
                        UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " +
              "stageid,starttime,status,remark,userid)  values ('" + publisherid + "','" + journalID + "','" + articleid + "'" +
              ",'STM1012',getdate(),'SUBMITTED','Article moved to Correction Queue as mail received.','caps1')");

                        UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " +
                    "stageid,starttime,status,remark,userid)  values ('" + publisherid + "','" + journalID + "','" + articleid + "'" +
                    ",'STM1013',getdate(),'STAGE-CHANGE','Article moved to Correction Queue as mail received.','caps1')");

                        UpdateQuery("update ldl_trn_AuthorEmailTracker set CorrectionReceived=getdate() where publisherid='"+publisherid+"' and journalid='"+journalID+"' and articleid='"+articleid+"' and category='author' and CorrectionReceived is null");

                        dt1 = GetDataFromDB("select t.stageid,d.email,t.publisherid,d.firstname,d.lastname from ldl_trn_articletracker t left join ldl_trn_authorDetails d on t.ARTICLEID=d.articleID and t.JOURNALID=d.journalid where d.IsCorrespondingAuthor='1' and foldercode='" + ArticleID + "'");

                        if (dt1.Rows.Count == 1)
                        {
                            dr1 = dt1.Rows[0];
                            string stageID = dr1[0].ToString();
                            string email = dr1[1].ToString();
                            string pid = dr1[2].ToString();
                            string fname = dr1[3].ToString();
                            string cupMailbBody = "Dear " + fname + ",<br><p>Thank you for submitting the corrections for " + journalID.ToUpper() + " " + articleid + ".</p><p>We will proceed with incorporating the corrections and will let you know if we have any queries as we progress.</p><br>Thank you<br><br>Journals Production<br>Lumina Datamatics";

                            string dbQuery = "insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('" + publisher + "','cup.caps@luminad.com','" + email + "','" + "cup.pm@luminad.com" + "','" + "" + journalID.ToUpper() + " " + articleid + " for correction" + "','" + cupMailbBody + "',getdate())";
                            UpdateQuery(dbQuery);

                            // blupencil
                            try
                            {
                                Console.WriteLine("Started for BluePencil");
                                //Process processChild = Process.Start(@"Z:\CAPS\bluePencil\bluePencil.bat", "publisherId::"+publisherid+" folderCode::"+ArticleID+" journalId::"+journalID+" articleId::"+articleid);

                                string argToPass = "-jar D:\\CAPS\\bluePencil\\bluePencil.jar publisherId::P1031 folderCode::" + ArticleID + " " + "journalId::" + journalID + " articleId::" + articleid;
                                ProcessStartInfo info = new ProcessStartInfo();
                                info.FileName = "java";
                                info.Arguments = argToPass;
                                info.CreateNoWindow = true;
                                info.UseShellExecute = true;
                                
                                try
                                {
                                    Process processChild = Process.Start(info);
                                    processChild.WaitForExit();
                                    processChild.Dispose();
                                }
                                catch(Exception e) {
                                    Console.WriteLine("BluPenil:"+ e.Message);
                                }

                                Console.WriteLine("End for BluPencil");
                            }
                            catch { }
                        }
                        result = true;
                        return true;
                    }
                    else
                    {
                        result = false;
                        return false;
                    }

                }
            }
            catch(Exception e) {
                result = false;
                Console.WriteLine(e.Message);
                return false;
            }
            return result;
        }

        private void updateForReSupply(string ArticleID, string dueDate, string TAT)
        {
            try
            {
                string[] articleDetails = ArticleID.Split('-');
                string aid = "";

                string query = "SELECT stageid FROM ldl_trn_articletracker where FOLDERCODE='" + ArticleID + "'";
                DataTable dt1 = GetDataFromDB(query);
                string stageID = "";
                if (dt1.Rows.Count == 1)
                {
                    DataRow dr1 = dt1.Rows[0];
                    stageID = dr1[0].ToString();
                }
                

                if (ArticleID.Contains("ANM"))
                {
                    try
                    {
                        aid = "-" + articleDetails[4];                        
                    }
                    catch(Exception e) {
                        aid = "";
                        Console.WriteLine(e.Message);
                    }
                }

                string dbQuery = "insert into ldl_trn_CUP_CAMSDetails (publisherid,journalid,articleid,downloadtime,deliverytype,CAMSDueDate,CalculatedCAMSDueDate) values ('P1031','" + articleDetails[2] + "','" + articleDetails[3] + aid + "',getdate(),'RE-SUPPLY','" + dueDate + "',dbo.fn_AddBusinessDays( getdate(),"+TAT+"))";
                UpdateQuery(dbQuery);

                if (dueDate.Trim() != "")
                {
                    UpdateQuery("Update ldl_trn_Articletracker set stageid='STM1079',CAMSDueDate='"+dueDate+"' where foldercode='" + ArticleID + "'");
                }
                else
                {
                    UpdateQuery("Update ldl_trn_Articletracker set stageid='STM1079' where foldercode='" + ArticleID + "'");
                }
                


                if (stageID != "")
                {
                    UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " +
   "stageid,starttime,status,remark,userid)  values ('P1031','" + articleDetails[2] + "','" + articleDetails[3]+aid + "'" +
   ",'" + stageID + "',getdate(),'SUBMITTED','','caps1')");

                }
                

                UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " +
            "stageid,starttime,status,remark,userid)  values ('P1031','" + articleDetails[2] + "','" + articleDetails[3]+aid + "'" +
            ",'STM1079',getdate(),'STAGE-CHANGE','Article moved to RESUPPLY as re-supply received.','caps1')");

            }
            catch { }
        }

        private void updateForSupply(string ArticleID, string dueDate, string stageid, string TAT, string moveTo, string artMetaStatus, string actionTaken)
        {
            try
            {
                
                string camsReceivedIn = "";                

                if (stageid == "STM1012")
                {
                    camsReceivedIn = " Author stage";
                    actionTaken = "No";
                }
                else if (stageid == "STM1032")
                {
                    camsReceivedIn = " PE stage";
                }

                string[] articleDetails = ArticleID.Split('-');
                string aid = "";
                if (ArticleID.Contains("ANM"))
                {
                    try
                    {

                        aid = "-"+articleDetails[4];
                    }
                    catch (Exception e)
                    {
                        aid = "";
                        Console.WriteLine(e.Message);
                    }
                }
                
                //dbo.fn_AddBusinessDays( getdate(),"+TAT+")
                string dbQuery = "insert into ldl_trn_CUP_CAMSDetails (publisherid,journalid,articleid,downloadtime,deliverytype,CAMSDueDate,CalculatedCAMSDueDate,ActionTaken) values ('P1031','" + articleDetails[2] + "','" + articleDetails[3] + aid + "',getdate(),'Supply','" + dueDate + "',dbo.UF_GetWorkingDay1(getdate()+" + TAT + "),'" + actionTaken + "')";
                UpdateQuery(dbQuery);

                if (dueDate.Trim() != "")
                {
                    UpdateQuery("Update ldl_trn_Articletracker set stageid='" + moveTo + "',CAMSDueDate='" + dueDate + "' where foldercode='" + ArticleID + "' and stageid='" + stageid + "'");
                }
                else
                {
                    UpdateQuery("Update ldl_trn_Articletracker set stageid='" + moveTo + "' where foldercode='" + ArticleID + "' and stageid ='" + stageid + "'");
                }                

                UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " +
   "stageid,starttime,status,remark,userid)  values ('P1031','" + articleDetails[2] + "','" + articleDetails[3]+aid + "'" +
   ",'"+stageid+"',getdate(),'SUBMITTED','','caps1')");

                UpdateQuery(" Insert into ldl_trn_ArticleMovementRegister_CUP (publisherid,journalid,articleid, " +
            "stageid,starttime,status,remark,userid)  values ('P1031','" + articleDetails[2] + "','" + articleDetails[3]+aid + "'" +
            ",'" + moveTo + "',getdate(),'STAGE-CHANGE','Article moved to FIRST VIEW as CAMS request received in " + camsReceivedIn + artMetaStatus + ".','caps1')");
            }
            catch { }
        }

        private string getDataByRegex(string mailbody, string regex)
        {
            string result = "";
            try
            {
                string pattern = regex;
                //if (fileName.Contains("asme") || fileName.Contains("seg"))


                //if (mailbody.Contains("ANM"))
                //{
                //    pattern = pattern + "(-)?[0-9]+";
                //}


                //List<string> articleID = new List<string>();
                Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
                MatchCollection matches = rgx.Matches(mailbody);

                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        result = match.Value;
                        break;
                    }
                }
                else
                {
                    matches = rgx.Matches(mailbody);
                    if (matches.Count > 0)
                    {
                        foreach (Match match in matches)
                        {
                            result = match.Value;
                            break;
                        }
                    }
                }
            }
            catch { }
            result = result.Replace("PII: ", "");
            return result;
        }


        private string getArticleID1(string msgSubject, string mailbody, string fileName, string mailID)
        {
            string articleID = "";
            try
            {
                string pattern = @"PII:.*?[0-9]+(X[0-9]+)?";
                //if (fileName.Contains("asme") || fileName.Contains("seg"))

                //List<string> articleID = new List<string>();
                Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
                MatchCollection matches = rgx.Matches(msgSubject);

                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        articleID = match.Value;
                        break;
                    }
                }
                else
                {
                    matches = rgx.Matches(mailbody);
                    if (matches.Count > 0)
                    {
                        foreach (Match match in matches)
                        {
                            articleID = match.Value;
                            break;
                        }
                    }
                }
            }
            catch { }
            articleID = articleID.Replace("PII: ", "");
            return articleID;
        }



        public void downloadZip(string host, string port, string user, string pass, string inPath, string ftpPath, string encryp,string dPath)
        {

            try
            {

                Console.WriteLine("inPath" + inPath);
                Console.WriteLine("ftpPath" + ftpPath);
                //string result="";
                //pass = "(n>2Uwc;?ZA`pqJ:";
               // string n = string.Format("{0:yyyy-MM-dd_hh-mm-ss-tt}", DateTime.Now);

                //String SuccessPath = ConfigurationManager.AppSettings["SuccessPath"].ToString() + "\\";
                //String InPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString() + "\\*.zip";
                //String FTPInPath = ConfigurationManager.AppSettings["FTPPath"].ToString();


                //String InPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString() + "\\*.zip";
                //String FTPInPath = ConfigurationManager.AppSettings["FTPPath"].ToString();

                //  1         2      3           4           5               6
                // HostName Port   UserName Password    InZipFilePath   FTPLocation

                String hostName = host;
                int portNumber = Convert.ToInt32(port);
                String userName = user;
                String password = pass;
                String InFilePath = inPath;
                String FTPInPath = ftpPath;
                bool EncryptionType = false;
                bool isSimple = false;

                try
                {
                    if (encryp.ToLower() == "explicit")
                    {
                        EncryptionType = true;
                    }
                    else if (encryp.ToLower() == "none")
                    {
                        isSimple = true;
                    }
                }
                catch (Exception e)
                {
                    isError = true;
                    Console.WriteLine(e.Message);
                }

                if (EncryptionType == false)
                {
                    // Setup session options
                    SessionOptions sessionOptions = null;
                    if (isSimple == false)
                    {
                        sessionOptions = new SessionOptions
                        {

                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit

                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.Implicit

                        };
                    }
                    else
                    {

                        sessionOptions = new SessionOptions
                        {
                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit

                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.None

                        };
                    }


                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        TransferOperationResult transferResult;

                        transferResult = session.GetFiles(FTPInPath, InFilePath, false, transferOptions);
                        //transferResult = session.PutFiles(InFilePath, FTPInPath, false, transferOptions);

                        //transferResult = session.PutFiles(InPath, "/users/J&JEditorial"+FTPInPath, false, transferOptions);
                        //transferResult = session.PutFiles(@"d:\toupload\*", "/LDL/SAP_testing/", false, transferOptions);

                        if (transferResult.IsSuccess == false)
                        {
                            session.Close();
                        }

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Download of {0} succeeded", transfer.FileName);
                            if (dPath.Contains("CUP"))
                            {
                                createLog(dPath, "Download of CAMS succeeded" + transfer.FileName);
                            }

                            //if (!System.IO.Directory.Exists(SuccessPath + "Success"))
                            //{
                            //    Directory.CreateDirectory(SuccessPath + "Success");
                            //}
                            //FileInfo finfo = new FileInfo(transfer.FileName);

                            //File.Move(transfer.FileName, SuccessPath + "Success\\" + n + "_" + finfo.Name);
                        }
                        session.Close();
                    }
                }

                else
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        //HostName = "vxandftps.astm.org",
                        //PortNumber=990,
                        //UserName = "J&JEditorial",
                        //Password = "8rUw$s@d",
                        //FtpSecure = FtpSecure.Implicit

                        //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                        HostName = hostName,
                        PortNumber = portNumber,
                        UserName = userName,
                        Password = password,
                        FtpSecure = FtpSecure.Explicit
                    };

                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);
                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        TransferOperationResult transferResult;

                        transferResult = session.PutFiles(InFilePath, FTPInPath, false, transferOptions);

                        if (transferResult.IsSuccess == false)
                        {
                            session.Close();
                        }

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                            if (dPath.Contains("CUP"))
                            {
                                createLog(dPath, "Download of CAMS succeeded" + transfer.FileName);
                            }
                        }
                        session.Close();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                isError = true;
                DataTable dt1 = GetDataFromDB("Select * from caps_trn_MailTracker where tomailid='rishiraj.sharma@luminad.com;sandeep.kumar@luminad.com' and mailsubject='" + e.Message.Replace("'", "''") + "'");
                if (dt1.Rows.Count ==0)
                {
                    UpdateQuery("INSERT INTO caps_trn_MailTracker (PublisherID, FromMailID, ToMailID, MailSubject, MailBody) VALUES ('P1031', 'cup.caps@luminad.com', 'rishiraj.sharma@luminad.com;sandeep.kumar@luminad.com', '" + e.Message.Replace("'", "") + "', '<p>Hi Team,</p><p>Please check for CAMS integeration fail.</p><p>Inpath: " + inPath + "</p><p>OutPath: " + ftpPath + "</p>  <p></p><p>Regards,</p><p>CAPSTeam</p>');");
                }
                
                
                createLog(dPath,"Error in package downloading: " + e.Message);
                //ErrorCAMS

                //Console.ReadLine();
            }
            finally { }
        }

        public void downloadZIPFromFTP(string host, string port, string user, string pass, string inPath, string ftpPath, string encryp)
        {

            try
            {

                Console.WriteLine("inPath" + inPath);
                Console.WriteLine("ftpPath" + ftpPath);

                //pass = "(n>2Uwc;?ZA`pqJ:";
                // string n = string.Format("{0:yyyy-MM-dd_hh-mm-ss-tt}", DateTime.Now);

                //String SuccessPath = ConfigurationManager.AppSettings["SuccessPath"].ToString() + "\\";
                //String InPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString() + "\\*.zip";
                //String FTPInPath = ConfigurationManager.AppSettings["FTPPath"].ToString();


                //String InPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString() + "\\*.zip";
                //String FTPInPath = ConfigurationManager.AppSettings["FTPPath"].ToString();

                //  1         2      3           4           5               6
                // HostName Port   UserName Password    InZipFilePath   FTPLocation

                String hostName = host;
                int portNumber = Convert.ToInt32(port);
                String userName = user;
                String password = pass;
                String InFilePath = inPath;
                String FTPInPath = ftpPath;
                bool EncryptionType = false;
                bool isSimple = false;

                try
                {
                    if (encryp.ToLower() == "explicit")
                    {
                        EncryptionType = true;
                    }
                    else if (encryp.ToLower() == "none")
                    {
                        isSimple = true;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                if (EncryptionType == false)
                {
                    // Setup session options
                    SessionOptions sessionOptions = null;
                    if (isSimple == false)
                    {
                        sessionOptions = new SessionOptions
                        {

                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit

                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.Implicit

                        };
                    }
                    else
                    {

                        sessionOptions = new SessionOptions
                        {
                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit

                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.None

                        };
                    }


                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        TransferOperationResult transferResult;

                        transferResult = session.GetFiles(FTPInPath, InFilePath, false, transferOptions);
                        //transferResult = session.PutFiles(InFilePath, FTPInPath, false, transferOptions);

                        //transferResult = session.PutFiles(InPath, "/users/J&JEditorial"+FTPInPath, false, transferOptions);
                        //transferResult = session.PutFiles(@"d:\toupload\*", "/LDL/SAP_testing/", false, transferOptions);

                        if (transferResult.IsSuccess == false)
                        {
                            session.Close();
                        }

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Download of {0} succeeded", transfer.FileName);

                            //if (!System.IO.Directory.Exists(SuccessPath + "Success"))
                            //{
                            //    Directory.CreateDirectory(SuccessPath + "Success");
                            //}
                            //FileInfo finfo = new FileInfo(transfer.FileName);

                            //File.Move(transfer.FileName, SuccessPath + "Success\\" + n + "_" + finfo.Name);
                        }
                        session.Close();
                    }
                }

                else
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        //HostName = "vxandftps.astm.org",
                        //PortNumber=990,
                        //UserName = "J&JEditorial",
                        //Password = "8rUw$s@d",
                        //FtpSecure = FtpSecure.Implicit

                        //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                        HostName = hostName,
                        PortNumber = portNumber,
                        UserName = userName,
                        Password = password,
                        FtpSecure = FtpSecure.Explicit
                    };

                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);
                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        TransferOperationResult transferResult;

                        transferResult = session.PutFiles(InFilePath, FTPInPath, false, transferOptions);

                        if (transferResult.IsSuccess == false)
                        {
                            session.Close();
                        }

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                        }
                        session.Close();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //Console.ReadLine();
            }
            finally { }
        }


        public Boolean uploadZIPToFTP_WileyCE(string host, string port, string user, string pass, string inPath, string ftpPath, string encryp)
        {
            bool fileCopied = false;
            try
            {

                //pass = "(n>2Uwc;?ZA`pqJ:";
                // string n = string.Format("{0:yyyy-MM-dd_hh-mm-ss-tt}", DateTime.Now);

                //String SuccessPath = ConfigurationManager.AppSettings["SuccessPath"].ToString() + "\\";
                //String InPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString() + "\\*.zip";
                //String FTPInPath = ConfigurationManager.AppSettings["FTPPath"].ToString();


                //String InPath = ConfigurationManager.AppSettings["DirectoryPath"].ToString() + "\\*.zip";
                //String FTPInPath = ConfigurationManager.AppSettings["FTPPath"].ToString();

                //  1         2      3           4           5               6
                // HostName Port   UserName Password    InZipFilePath   FTPLocation

                String hostName = host;
                int portNumber = Convert.ToInt32(port);
                String userName = user;
                String password = pass;
                String InFilePath = inPath;
                String FTPInPath = ftpPath;
                bool EncryptionType = false;
                bool isSimple = false;

                try
                {
                    if (encryp.ToLower() == "explicit")
                    {
                        EncryptionType = true;
                    }
                    else if (encryp.ToLower() == "none")
                    {
                        isSimple = true;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                if (EncryptionType == false)
                {
                    // Setup session options
                    SessionOptions sessionOptions = null;
                    if (isSimple == false)
                    {
                        sessionOptions = new SessionOptions
                        {

                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit

                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.Implicit

                        };
                    }
                    else
                    {
                        sessionOptions = new SessionOptions
                        {
                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit
                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.None

                        };
                    }

                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        //TransferOperationResult transferResult;

                        session.MoveFile(FTPInPath, InFilePath);
                        session.MoveFile(FTPInPath.Replace(".zip", ".go.xml"), InFilePath.Replace(".zip", ".go.xml"));

                        if (session.FileExists(InFilePath))
                        {
                            fileCopied = true;
                        }
                        //transferResult = session.GetFiles(FTPInPath, InFilePath, false, transferOptions);


                        //transferResult = session.PutFiles(InFilePath, FTPInPath, false, transferOptions);
                        //transferResult = session.PutFiles(InPath, "/users/J&JEditorial"+FTPInPath, false, transferOptions);
                        //transferResult = session.PutFiles(@"d:\toupload\*", "/LDL/SAP_testing/", false, transferOptions);

                        //if (transferResult.IsSuccess == false)
                        //{
                        //    session.Close();
                        //}

                        //// Throw on any error
                        //transferResult.Check();

                        //// Print results
                        //foreach (TransferEventArgs transfer in transferResult.Transfers)
                        //{
                        //    Console.WriteLine("Upload of {0} succeeded", transfer.FileName);

                        //    //if (!System.IO.Directory.Exists(SuccessPath + "Success"))
                        //    //{
                        //    //    Directory.CreateDirectory(SuccessPath + "Success");
                        //    //}
                        //    //FileInfo finfo = new FileInfo(transfer.FileName);

                        //    //File.Move(transfer.FileName, SuccessPath + "Success\\" + n + "_" + finfo.Name);
                        //}
                        session.Close();
                    }
                }

                else
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        //HostName = "vxandftps.astm.org",
                        //PortNumber=990,
                        //UserName = "J&JEditorial",
                        //Password = "8rUw$s@d",
                        //FtpSecure = FtpSecure.Implicit

                        //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                        HostName = hostName,
                        PortNumber = portNumber,
                        UserName = userName,
                        Password = password,
                        FtpSecure = FtpSecure.Explicit

                    };

                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);
                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        TransferOperationResult transferResult;

                        transferResult = session.PutFiles(InFilePath, FTPInPath, false, transferOptions);

                        if (transferResult.IsSuccess == false)
                        {
                            session.Close();
                        }

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                        }
                        session.Close();
                    }
                }


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //Console.ReadLine();
                isError = true;
            }
            finally {  }
            return fileCopied;
        }


        public Boolean moveProcessedFile(string host, string port, string user, string pass, string inPath, string ftpPath, string encryp)
        {
            bool fileCopied = false;
            try
            {

                String hostName = host;
                int portNumber = Convert.ToInt32(port);
                String userName = user;
                String password = pass;
                String InFilePath = inPath;
                String FTPInPath = ftpPath;
                bool EncryptionType = false;
                bool isSimple = false;

                try
                {
                    if (encryp.ToLower() == "explicit")
                    {
                        EncryptionType = true;
                    }
                    else if (encryp.ToLower() == "none")
                    {
                        isSimple = true;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                if (EncryptionType == false)
                {
                    // Setup session options
                    SessionOptions sessionOptions = null;
                    if (isSimple == false)
                    {
                        sessionOptions = new SessionOptions
                        {

                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit

                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.Implicit

                        };
                    }
                    else
                    {
                        sessionOptions = new SessionOptions
                        {
                            Protocol = Protocol.Ftp,
                            //HostName = "vxandftps.astm.org",
                            //PortNumber=990,
                            //UserName = "J&JEditorial",
                            //Password = "8rUw$s@d",
                            //FtpSecure = FtpSecure.Implicit
                            //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                            HostName = hostName,
                            PortNumber = portNumber,
                            UserName = userName,
                            Password = password,
                            FtpSecure = FtpSecure.None

                        };
                    }

                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        //TransferOperationResult transferResult;

                        session.MoveFile(FTPInPath, InFilePath);                       

                        if (session.FileExists(InFilePath))
                        {
                            fileCopied = true;
                        }
                        
                        session.Close();
                    }
                }

                else
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        //HostName = "vxandftps.astm.org",
                        //PortNumber=990,
                        //UserName = "J&JEditorial",
                        //Password = "8rUw$s@d",
                        //FtpSecure = FtpSecure.Implicit

                        //,SshHostKeyFingerprint = "ssh-rsa 1024 3a:87:bc:d9:f2:27:2a:a5:f6:a0:6f:a8:ba:dc:9c:1b"

                        HostName = hostName,
                        PortNumber = portNumber,
                        UserName = userName,
                        Password = password,
                        FtpSecure = FtpSecure.Explicit

                    };

                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);
                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                        //null for default permissions.  Can set user, 
                        //Group, or other Read/Write/Execute permissions. 
                        transferOptions.PreserveTimestamp = false; //Set last write time of 
                        //destination file to that of source file - basically change the timestamp 
                        //to match destination and source files.   
                        //transferOptions.ResmeSupport.State = TransferResumeSupportState.Off;

                        TransferOperationResult transferResult;

                        transferResult = session.PutFiles(InFilePath, FTPInPath, false, transferOptions);

                        if (transferResult.IsSuccess == false)
                        {
                            session.Close();
                        }

                        // Throw on any error
                        transferResult.Check();

                        // Print results
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                        }
                        session.Close();
                    }
                }


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //Console.ReadLine();
                isError = true;
            }
            finally { }
            return fileCopied;
        }


        private void mailSaveToFTP(string Publisher,string fileName, string mailbody)
        {
            try
            {
                string filePath = "";
                if (Publisher == "CSP")
                {
                    filePath = ConfigurationManager.AppSettings["CSPIntegerationPath"].ToString();
                }
                //StreamWriter sw = new StreamWriter(filePath + "\\" + fileName+".txt", true, Encoding.UTF8);
                StreamWriter sw = new StreamWriter(filePath + "\\" + fileName + ".txt", true, Encoding.ASCII);
                sw.WriteLine(mailbody);
                sw.Close();
            }
            catch { }


        }



        public static void SendEmail(string subject, string body,string formMailID,string passwrod,string publisherid)
        {
            try
            {
                System.Net.NetworkCredential nc = new System.Net.NetworkCredential(formMailID, passwrod);
                string server = "mail.luminad.com";
                string to = "Melissa.metz@luminad.com,";
                //string to = "rishiraj.sharma@luminad.com,";
                if (publisherid.Contains("SPIE"))
                {
                    to = to + "spie-manager@luminad.com";
                    //to = to + "sandeep.kumar@luminad.com";
                }
                else
                {
                    to = to.TrimEnd(',');
                }
                //string to = "rishiraj.sharma@luminad.com";
                string bcc = "rishiraj.sharma@luminad.com";
                string from = formMailID;
                //MailAddress add=new MailAddress(
                MailMessage message = new MailMessage(from, to);
                message.Bcc.Add(bcc);
                message.Subject = subject;
                message.Body = body;
                message.IsBodyHtml = true;
                SmtpClient client = new SmtpClient(server, 25);
                client.UseDefaultCredentials = false;
                client.Credentials = nc;
                client.EnableSsl = true;

                try
                {
                    //client.Send(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught in SendMail(): {0}",
                          ex.ToString() + " : " + subject + " : Details:: " + body);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + " : " + subject + " : Details:: " + body);
            }
        }




        public static void SendEmail_WithToMailID(string subject, string body, string formMailID, string passwrod, string publisherid,String toMailID,string ccMailID)
        {
            try
            {
                System.Net.NetworkCredential nc = new System.Net.NetworkCredential(formMailID, passwrod);
                string server = "mail.luminad.com";
                //string to = "rishiraj.sharma@luminad.com";
                //ccMailID = "rishiraj.sharma@luminad.com";

                string to = toMailID;
                string bcc = "rishiraj.sharma@luminad.com";
                string from = formMailID;
                //MailAddress add=new MailAddress(
                MailMessage message = new MailMessage(from, to);

                if (ccMailID.Trim() != "")
                {
                    message.CC.Add(ccMailID);
                }

                message.Bcc.Add(bcc);
                message.Subject = subject;
                message.Body = body;
                message.IsBodyHtml = true;
                SmtpClient client = new SmtpClient(server, 25);
                client.UseDefaultCredentials = false;
                client.Credentials = nc;
                client.EnableSsl = true;

                try
                {
                    //client.Send(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught in SendMail(): {0}",
                          ex.ToString() + " : " + subject + " : Details:: " + body);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + " : " + subject + " : Details:: " + body);
            }
        }

        public string UpdateQuery(string dbQuery)
        {
            string retrunVal = "";
            SqlConnection cnn;
            string conStr = ConfigurationManager.AppSettings["MyKey"].ToString();
            string connectionString = "";
            connectionString = conStr;
            cnn = new SqlConnection(connectionString);
            cnn.Open();

            try
            {
                //string strQuery = "Update ldl_trn_articletracker set STAGEID='STM1006'  where FOLDERCODE ='" + fileN + "'";
                string strQuery = dbQuery;
                SqlCommand cmd = new SqlCommand(dbQuery, cnn);
                cmd.ExecuteNonQuery();
                //Console.WriteLine(str + " " + fileN);
                cnn.Close();
                cnn = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("DataManupluation----------" + e.Message + "  dbQuery:: " + dbQuery);
                cnn.Close();
                cnn = null;
                try {
                    string path = @"D:\CAPS\UpdateQuery.txt";
                    if (!File.Exists(path))
                    {
                        // Create a file to write to.
                        using (StreamWriter sw = File.CreateText(path))
                        {
                            sw.WriteLine(dbQuery);
                        }
                    }
                    else
                    {
                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine(dbQuery);
                        }
                    }
                }
                catch { }
                
                
            }
            return retrunVal;
        }


        public DataTable GetDataFromDB(string query)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection cnn;
                string conStr = ConfigurationManager.AppSettings["MyKey"].ToString();
                string connectionString = "";
                connectionString = conStr;
                cnn = new SqlConnection(connectionString);
                cnn.Open();
                try
                {
                    SqlCommand cmd = new SqlCommand(query, cnn);
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandTimeout = 3000;
                    // Create a DataAdapter to run the command and fill the DataTable
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = cmd;
                    da.Fill(dt);
                    cnn.Close();
                }
                catch
                {
                    //Console.WriteLine("Step sql conn");
                    cnn.Close();

                    try
                    {
                        string path = @"D:\CAPS\UpdateQuery.txt";
                        if (!File.Exists(path))
                        {
                            // Create a file to write to.
                            using (StreamWriter sw = File.CreateText(path))
                            {
                                sw.WriteLine(query);
                            }
                        }
                        else
                        {
                            using (StreamWriter sw = File.AppendText(path))
                            {
                                sw.WriteLine(query);
                            }
                        }
                    }
                    catch { }

                }
            }
            catch
            {
                //Console.WriteLine("Step 123");
            }
            return dt;
        }

        // Getting Email Sender Name
        private string getSenderName(Message message)
        {
            string recFrom;
            try
            {
                recFrom = message.Headers.From.Address;                
                if (recFrom == "")
                {
                    recFrom = "IDNotFound";
                }
                else
                {
                    string[] emailArr = recFrom.Split('@');
                    recFrom = emailArr[0];
                }
            }
            catch
            {
                recFrom = "IDNotFound";
            }
            return recFrom;
        }

        // Getting Message Subject 
        private string getMessageSubject(Message message)
        {
            string msgSubject = "No subject";
            try
            {
                msgSubject = CleanInput(message.Headers.Subject);
                
                if (msgSubject == "")
                {
                    msgSubject = "No subject";
                }
            }
            catch
            {
                msgSubject = "No subject";
            }
            return msgSubject;
        }

        // Removing characters--- those are not alphabetes and .@-
        static string CleanInput(string strIn)
        {
            // Replace invalid characters with empty strings. 
            try
            {
                strIn = strIn.Replace(":", ": ");
                strIn = strIn.Replace(":  ", ": ");
                //return Regex.Replace(strIn, @"[^\w\s\.@-]", "",
                //                     RegexOptions.None, TimeSpan.FromSeconds(1.5));
                return Regex.Replace(strIn, @"[^\w\s\.@-]", "",RegexOptions.None);
            }
            // If we timeout when replacing invalid characters,  
            // we should return Empty. 
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return String.Empty;
            }
        }

        // Getting article Id by regex.
        private string getArticleID(string msgSubject, string mailbody, string fileName, string mailID)
        {
            string articleID = "";
            try
            {
                string pattern = @"([A-Z]+)([0-9]+)?(\-)([0-9]+([A-Z]+)?)";


                if (fileName.Contains("astm") || mailID.Contains("astm"))
                {
                    pattern = @"([A-Z]+)(\.)?(\-)?([0-9]+)(\-)([0-9]+)";
                }

                if (fileName.Contains("NACE") || mailID.Contains("nace"))
                {
                    pattern = @"([A-Z]+)(\.)?(\-)?([0-9]+)(\-)([0-9]+)";
                }
                if (mailID.Contains("asme") || mailID.Contains("csp") || mailID.Contains("seg") || fileName.Contains("asme") || fileName.Contains("seg"))
                {
                    pattern = @"([A-Z]+)([0-9]+)?(\-)([0-9]+([A-Z]+)?)(\-)([0-9]+([A-Z]+)?)";
                }

                if (fileName.Contains("aiaa") || mailID.Contains("aiaa") )
                {
                    pattern = @"(([A-Z]+)([0-9]+))?(\-?)(([A-Z]+[0-9]+))";
                }

                if (mailID.Contains("hk") || fileName.Contains("hk") || mailID.Contains("aoac") ||  mailID.Contains("wiley"))
                {
                    pattern = @"([A-Z]+)([0-9]+)?(\.)?(\-)?([0-9]+([A-Z]+)?)(\-)([0-9]+([A-Z]+)?)";
                }

                if (mailID.Contains("wiley") && msgSubject.Contains("for ") && msgSubject.Contains("- EMID"))
                {
                    pattern = "for [A-Z]+(\\.)?[0-9]+";
                }
                else if(fileName.Contains("wiley") || mailID.Contains("wiley"))
                {
                    msgSubject=msgSubject.Replace("R1","").Replace("R2","");
                    pattern = @"([A-Z]+)([0-9]+)?(\.)?(\-)?([0-9]+([A-Z]+)?)";
                }


                if (mailID.Contains("spie") && msgSubject.Contains(" AP-"))
                {
                    pattern = @"([A-Z]+)([0-9]+)?(\-)([0-9]+([A-Z]+)?)((\-)([0-9]+([A-Z]+)?))?";
                }

                //string pattern = @"([A-Z]+)([0-9]+)?(\-)([0-9]+([A-Z]+)?)";
                
                //List<string> articleID = new List<string>();
                Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
                MatchCollection matches = rgx.Matches(msgSubject);

                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        articleID = match.Value;
                        break;
                    }
                }
                else
                {
                    matches = rgx.Matches(mailbody);

                    if (matches.Count > 0)
                    {
                        foreach (Match match in matches)
                        {
                            articleID = match.Value;
                            break;
                        }
                    }
                }
            }
            catch { }

            if (articleID.Contains("for") && mailID.Contains("wiley") && msgSubject.Contains("for ") && msgSubject.Contains("- EMID"))
            {
                articleID= articleID.Replace("for", "");
                articleID = articleID.Replace(" ", "");
                string jid = getDataByRegex(articleID, "[A-Z]+");
                string aid = getDataByRegex(articleID, "[0-9]+");
                articleID = jid.ToUpper() + "-" + aid;
            }
            return articleID;
        }


        public void createLog(string dPath,string msg)
        {
            try
            {
                if (dPath.ToLower().Contains("cup"))
                {
                    DateTime now = DateTime.Now;
                    Console.WriteLine(now);

                    if (!Directory.Exists(dPath + "\\OQC-LOG"))
                    {
                        Directory.CreateDirectory(dPath + "\\OQC-LOG");
                    }
                    //if (!File.Exists(dPath + "\\OQC-LOG\\LOG.txt"))
                    //{
                    //    File.Create(dPath + "\\OQC-LOG\\LOG.txt");
                    //}

                    using (StreamWriter sr = File.AppendText(dPath + "\\OQC-LOG\\LOG.txt"))
                    {                        
                        sr.WriteLine(msg);
                        sr.Close();
                    }
                }
            }
            catch(Exception e) {
                Console.WriteLine(e.Message);
            }
           
        }


        public DateTime timeZoneConverter(DateTime dateTime1)
        {
            TimeZoneInfo timeZoneInfo;
            DateTime dateTime;
            //Set the time zone information to India Standard Time 
            timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
            //Get date and time in India Standard Time 
            dateTime = TimeZoneInfo.ConvertTime(DateTime.Now, timeZoneInfo);
            //Print out the date and time
            //Console.WriteLine(dateTime.ToString("dddd, dd MMMM yyyy HH:mm:ss"));
            return dateTime1;
        }
    }
}
//try
//{
//    FileInfo finfo1 = new FileInfo(emlPath);
//    message.Save(finfo1);
//    //Message loadedMessage = Message.Load(finfo1);
//}
//catch { }

//}

 //if (publisher == "P1031" &&  fromMailID.ToLower().Contains("cams@cambridge.org"))
 //                                       {
 //                                           //Thu, 13 Jun 2019 12:25:24 +0000
 //                                           //string timeReceived = "";
 //                                           try
 //                                           {
 //                                               //string[] timeReceivedArr = recTime.Split('+');
 //                                               //DateTime timeIST = timeZoneConverter(Convert.ToDateTime(timeReceivedArr[0].Trim()));
 //                                               //timeReceived = " Sent: (UK Timing): " + timeReceivedArr[0].Trim() + "   (IST): " + timeIST.ToString();

 //                                           }
 //                                           catch { }

 //                                           //UpdateQuery("insert into  caps_trn_MailTracker (publisherid,frommailid,tomailid,ccmailid,mailsubject,mailbody,entrytime) values ('P1031','cup.caps@luminad.com','cupcams.pdy@luminad.com','capssupport@luminad.com','" + msgSubject.Replace("'", "''") + "','" + "<p>" + timeReceived + "</p><p>This mail is forwarded by CAPS Team. </p><br/>" + mb + "',getdate())");
 //                                       }
