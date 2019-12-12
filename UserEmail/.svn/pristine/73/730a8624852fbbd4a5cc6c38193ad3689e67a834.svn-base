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
namespace UserEmails
{
    public class GetUserEmail
    {
        public readonly Dictionary<int, Message> messages = new Dictionary<int, Message>();
        public readonly Pop3Client pop3Client;
        private TreeView listMessages;

        public GetUserEmail()
        {
            pop3Client = new Pop3Client();
        }
        public void ReceiveMails()
        {
            // Disable buttons while working

            try
            {

                if (pop3Client.Connected)
                    pop3Client.Disconnect();
                pop3Client.Connect("mail.luminad.com", 995, true);
                pop3Client.Authenticate("rishiraj.sharma@luminad.com", "59D6UaN2");
                

                int count = pop3Client.GetMessageCount();
                

                
                messages.Clear();
                
                
                int success = 0;
                int fail = 0;
                for (int i = count; i >= 1; i -= 1)
                {
                    try
                    {
                        Message message;
                        try
                        {
                            message = pop3Client.GetMessage(i);
                        }
                        catch
                        {
                            return;
                        }

                        
                        messages.Add(i, message);

                        string recFrom = "";
                        try
                        {
                            recFrom = message.Headers.From.Address;
                        }
                        catch
                        {
                            recFrom = "IDNotFound";
                        }

                        if (recFrom == "")
                        {
                            recFrom = "IDNotFound";
                        }
                        else
                        {
                            string[] emailArr = recFrom.Split('@');
                            recFrom = emailArr[0];
                        }

                        //27 Jan 2015 12:40:11 +0530
                        string recTime = message.Headers.Date;

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
                        string longsubject = "";
                        bool longsubjectBool = false;
                        string msgSubject ="";

                        try
                        {
                            msgSubject = CleanInput(message.Headers.Subject);
                            if (msgSubject.Length > 150)
                            {
                                longsubject = msgSubject;
                                msgSubject = msgSubject.Substring(0, 50);
                                msgSubject = msgSubject + "....";
                                longsubjectBool = true;
                            }
                            if(msgSubject=="")
                            {
                                msgSubject = "No subject";
                            }
                        }
                        catch
                        {
                            msgSubject = "No subject";
                        }

                        string fileToSave = recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + recTimeArr[4].Replace(':','.') + "_ldps_" +  msgSubject;
                        string fileName = recFrom;
                        string dirPath = @"d:\rishi\email\" + fileName + @"\";

                        if (!Directory.Exists(dirPath))
                        {
                            Directory.CreateDirectory(dirPath);
                        }
                        System.Net.Mail.MailMessage mailmsg = message.ToMailMessage();
                        mailmsg.IsBodyHtml = true;
                        string mailbody = mailmsg.Body;
                        MessagePart msgBody = message.FindFirstPlainTextVersion();

                        if (msgBody != null)
                        {
                            mailbody = message.FindFirstPlainTextVersion().GetBodyAsText();
                        }

                        
                        string msgAttachment = "";                        
                        string attTime = recTimeArr[1] + recTimeArr[2] + recTimeArr[3] + "_ldps_" + recTimeArr[4].Replace(':', '.')+ "_ldps_";
                        foreach (MessagePart mpart in message.FindAllAttachments())
                        {
                            msgAttachment = "ldpsattachement";
                            string fileNameAttach = mpart.FileName;                            
                            string file_name_attach = mpart.ContentType.MediaType;
                            FileInfo finfo = new FileInfo(dirPath + attTime + fileNameAttach);
                            mpart.Save(finfo);
                        }
                        string path = "";
                        if (msgAttachment != "")
                        {
                            //path = dirPath + fileToSave + "_" + msgAttachment  + msgFiles.TrimEnd('_') + ".mail";
                            //path = dirPath + fileToSave + "_ldps_" + msgAttachment + ".mail";
                            path = dirPath + fileToSave + ".mail";
                        }
                        else
                        {
                            path = dirPath + fileToSave + ".mail";
                        }

                        if (!File.Exists(path))
                        {                            
                            using (StreamWriter sw = File.CreateText(path))
                            {
                                if(longsubjectBool==true)
                                {
                                    sw.WriteLine("Subject: " +  longsubject + "\n");
                                }
                                sw.WriteLine(mailbody);
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


                ////MessageBox.Show(this, "Mail received!\nSuccesses: " + success + "\nFailed: " + fail, "Message fetching done");

                if (fail > 0)
                {
                    //MessageBox.Show(this,
                    //"Since some of the emails were not parsed correctly (exceptions were thrown)\r\n" +
                    //"please consider sending your log file to the developer for fixing.\r\n" +
                    //"If you are able to include any extra information, please do so.",
                    //"Help improve OpenPop!");
                }
            }
            catch (InvalidLoginException) { }
            catch (PopServerNotFoundException) { }
            catch (PopServerLockedException){}
            catch (LoginDelayException){}            
            catch {}
            finally
            {                
            }
        }

        static string CleanInput(string strIn)
        {
            // Replace invalid characters with empty strings. 
            try
            {
                return Regex.Replace(strIn, @"[^\w\s\.@-]", "",
                                     RegexOptions.None, TimeSpan.FromSeconds(1.5));
            }
            // If we timeout when replacing invalid characters,  
            // we should return Empty. 
            catch (RegexMatchTimeoutException)
            {
                return String.Empty;
            }
        }
    }
}
