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
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Data;


namespace UserEmails
{
    class Program
    {
        
        static void Main(string[] args)
        {
            
            GetUserEmail pg = new GetUserEmail();

        //pg.updateForCAMSReceived("LDL-CUP-BAJ-1900008", @"E:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-CUP-BAJ-1900008\", "BAJ", "1900008");
        //pg.updateForCAMSReceived("LDL-CUP-NJG-1900001", @"E:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-CUP-NJG-1900001\", "NJG", "1900001");
        //pg.updateForCAMSReceived("LDL-CUP-NJG-1900008", @"E:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-CUP-NJG-1900008\", "NJG", "1900008");
        //pg.updateForCAMSReceived("LDL-CUP-PHC-1900061", @"E:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-CUP-PHC-1900061\", "PHC", "1900061");
        //pg.updateForCAMSReceived("LDL-CUP-PHC-1900072", @"E:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-CUP-PHC-1900072\", "PHC", "1900072");



        chkAgain:
            GetUserEmail euMail = new GetUserEmail();
            euMail.ReceiveMails();
            System.Threading.Thread.Sleep(80000);
            goto chkAgain;
        }


        public void test1()
        {
            //GetUserEmail gu = new GetUserEmail();
            //gu.versioningOfFile(@"D:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-WILEY-ENTE-201900928" + "\\RT_Proof_Correction_Processing", @"D:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-WILEY-ENTE-201900928" + "\\" + "ente" + "_" + "326ca4e8-95dc-4c4d-a14e-fa6f9c8634a6" + ".zip", ArticleID);

        }
        public static  void runTest()
        {
            try
            {
                Console.WriteLine("Started for BluePencil");
                //Process processChild = Process.Start(@"Z:\CAPS\bluePencil\bluePencil.bat", "publisherId::"+publisherid+" folderCode::"+ArticleID+" journalId::"+journalID+" articleId::"+articleid);

                string argToPass = "-jar D:\\CAPS\\bluePencil\\bluePencil.jar publisherId::P1031 folderCode::" + "aid" + " " + "journalId::" + "jid" + " articleId::" + "aid";
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
                catch (Exception e)
                {
                    Console.WriteLine("BluPenil:" + e.Message);
                }

                Console.WriteLine("End for BluPencil");
            }
            catch(Exception e) {
                Console.WriteLine("excep:"+e.Message);
            }
            Console.ReadLine();
        }
        public void test(string bundleType)
        {


            string ArticleID = "LDL-WILEY-ENTE-201900928";
            string dirStartPath = @"D:\server\tomcat8\webapps\ROOT\stockage\stock\";
            string[] ard = "LDL-WILEY-ENTE-201900928".Split('-');
            Console.WriteLine(ard.Length);
            GetUserEmail gu = new GetUserEmail();

            if (bundleType.ToLower()=="ce")
            {
                gu.removeTempDirAndFiles(dirStartPath + ArticleID + "\\temp");
                bool result = gu.extractFiles(@"D:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-WILEY-ENTE-201900928\", @"D:\server\tomcat8\webapps\ROOT\stockage\stock\LDL-WILEY-ENTE-201900928\ente_d0c7c41a-5457-48ab-9aa4-e02f13e33d67.zip", "LDL-WILEY-ENTE-201900928");

                string bundlePath = dirStartPath + ArticleID + "\\ente_d0c7c41a-5457-48ab-9aa4-e02f13e33d67.zip";
                if (result == true)
                {
                    bundlePath = dirStartPath + ArticleID + "\\temp\\ente_d0c7c41a-5457-48ab-9aa4-e02f13e33d67.zip";
                }

                //CEBundle AABundle
                gu.versioningofDownloadFiles(@"D:\server\tomcat8\webapps\ROOT\stockage\stock\" + "LDL-WILEY-ENTE-201900928" + "\\", "JOB ANALYSIS", ard[2] + "_" + ard[3], "CEBundle", bundlePath, ard[2] + "_" + ard[3] + "_JA_Input.zip", true);

                if (result == true)
                {
                    gu.removeTempDirAndFiles(dirStartPath + ArticleID + "\\temp");
                }

                FileInfo f = new FileInfo(bundlePath);
                gu.removeTempDirAndFiles(bundlePath);

                if (Directory.Exists(dirStartPath + ArticleID + "\\ente_d0c7c41a-5457-48ab-9aa4-e02f13e33d67"))
                {
                    gu.removeTempDirAndFiles(dirStartPath + ArticleID + "\\ente_d0c7c41a-5457-48ab-9aa4-e02f13e33d67");
                }

                 DataTable dtHoldDetails = gu.GetDataFromDB("select foldercode,stageid,ISONHOLD from ldl_trn_articletracker where foldercode='" + ArticleID + "' ");

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
                        gu.UpdateQuery("Update ldl_trn_articletracker set isonhold='N' where foldercode='" + ArticleID + "';");
                        gu.UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'" + stage + "','STAGE-CHANGE','Hold removed, mail received from wiley as: CE integeration','caps1');");
                        gu.UpdateQuery("Update ldl_trn_ArticleHoldDetails set unholdtime=getdate(),unholdstatus='UNHOLD',unholduserid='caps1',unholdcomment='Hold removed, mail received from wiley as: CE integeration' where journalid='" + ard[2] + "' and articleid='" + ard[3] + "' and publisherid='P1032' and unholdstatus is null;");
                    }
                 }
                gu.UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'" + stage  + "','','Mail received from wiley as: CE integeration','caps1');");                

            }else if (bundleType.ToLower() == "aa")
            {
                //CEBundle AABundle
                string bundlePath = dirStartPath + ArticleID + "\\ente_d0c7c41a-5457-48ab-9aa4-e02f13e33d67.zip";
                gu.versioningofDownloadFiles(@"D:\server\tomcat8\webapps\ROOT\stockage\stock\" + "LDL-WILEY-ENTE-201900928" + "\\", "JOB ANALYSIS", ard[2] + "_" + ard[3], "AABundle", bundlePath, ard[2] + "_" + ard[3] + "_JA_Input.zip", true);

                DataTable dtHoldDetails = gu.GetDataFromDB("select foldercode,stageid,ISONHOLD from ldl_trn_articletracker where foldercode='" + ArticleID + "' ");

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
                        gu.UpdateQuery("Update ldl_trn_articletracker set isonhold='N' where foldercode='" + ArticleID + "';");
                        gu.UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'STM1023','STAGE-CHANGE','Hold removed, mail received from wiley as: AA integeration','caps1');");
                        gu.UpdateQuery("Update ldl_trn_ArticleHoldDetails set unholdtime=getdate(), unholdstatus='UNHOLD',unholduserid='caps1',unholdcomment='Hold removed, mail received from wiley as: AA integeration' where journalid='" + ard[2] + "' and articleid='" + ard[3] + "' and publisherid='P1032' and unholdstatus is null;");
                    }
                }
                gu.UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + ard[2] + "','" + ard[3] + "',getdate(),'"+stage+"','','Mail received from wiley as: AA integeration','caps1');");
            }
            else if (bundleType.ToLower() == "aac")
            {
                //AA task has been cancelled
                string[] folderCode = ArticleID.Split('-');
                gu.UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1023','STAGE-CHANGE','Mail received from wiley as: AA task has been cancelled','caps1');");
                gu.UpdateQuery("Update ldl_trn_articletracker set stageid='STM1023' where foldercode='" + ArticleID + "';Update ldl_trn_Articledetails set IsWileyReassigned='Yes' where journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "' and publisherid='P1032'");
            }
            else if (bundleType.ToLower() == "cc")
            {
                // CE task has been cancelled in CE stage
                string[] folderCode = ArticleID.Split('-');
                gu.UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1002','STAGE-CHANGE','Mail received from wiley as: CE task has been cancelled in CE stage','caps1');");
                gu.UpdateQuery("Update ldl_trn_articletracker set stageid='STM1002' where foldercode='" + ArticleID + " and stageid='STM1003''; Update ldl_trn_Articledetails set IsWileyReassigned='Yes' where journalid='"+folderCode[2]+"' and articleid='"+folderCode[3]+"' and publisherid='P1032'");
            }
            else if (bundleType.ToLower() == "cc1")
            {
                // CE task has been cancelled in non CE stage
                string[] folderCode = ArticleID.Split('-');
                gu.UpdateQuery("insert into ldl_trn_ArticleMovementRegister_WILEY (publisherid,journalid,articleid,starttime,stageid,status,remark,userid) values ('P1032','" + folderCode[2] + "','" + folderCode[3] + "',getdate(),'STM1023','STAGE-CHANGE','Mail received from wiley as: CE task has been cancelled','caps1');");
                //gu.UpdateQuery("Update ldl_trn_articletracker set stageid='STM1002' where foldercode='" + ArticleID + " and stageid='STM1003''; Update ldl_trn_Articledetails set IsWileyReassigned='Yes' where journalid='" + folderCode[2] + "' and articleid='" + folderCode[3] + "' and publisherid='P1032'");
            }      
        }

    }
}








