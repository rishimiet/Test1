using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Apache.NMS;
using Apache.NMS.Util;
using Apache.NMS.ActiveMQ;
using System.Threading;
using System.Configuration;
namespace UserEmails
{
    class JMSProducer
    {

        public void postMessage(string txt,string msgg)
        {
            //Create the Connection factory
            try
            {
                string url = ConfigurationManager.AppSettings["queueURL"].ToString();
                string queueName = ConfigurationManager.AppSettings["queueName"].ToString();

                if (msgg == "Redline")
                {
                    queueName = ConfigurationManager.AppSettings["queueRedline"].ToString();
                }
                IConnectionFactory factory = new ConnectionFactory(url);

                using (IConnection connection = factory.CreateConnection())
                using (ISession session = connection.CreateSession())
                {
                    IDestination destination = SessionUtil.GetDestination(session, queueName);
                    Console.WriteLine("Using destination: " + destination);

                    // Create a producer                
                    using (IMessageProducer producer = session.CreateProducer(destination))
                    {
                        // Start the connection so that messages will be processed.
                        connection.Start();
                        IMapMessage msg = producer.CreateMapMessage();
                        string fileName = txt;
                        string[] arrStr = txt.Split(',');

                        msg.Body.SetString("FILENAME", txt);
                        if (msgg == "Redline")
                        {
                            msg.Body.SetString("MESSAGE", "");
                        }
                        //msg.Body.SetString("MESSAGE", msgg);
                        producer.Send(msg);
                        Console.WriteLine("Message Send for " + txt);
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }

        public void postMessageforCompareMeta(string txt, string msgg)
        {
            //Create the Connection factory
            try
            {
                string url = ConfigurationManager.AppSettings["queueURL"].ToString();
                string queueName = ConfigurationManager.AppSettings["queueName"].ToString();

                IConnectionFactory factory = new ConnectionFactory(url);

                using (IConnection connection = factory.CreateConnection())
                using (ISession session = connection.CreateSession())
                {
                    IDestination destination = SessionUtil.GetDestination(session, queueName);
                    Console.WriteLine("Using destination: " + destination);

                    // Create a producer                
                    using (IMessageProducer producer = session.CreateProducer(destination))
                    {
                        // Start the connection so that messages will be processed.
                        connection.Start();
                        IMapMessage msg = producer.CreateMapMessage();
                        string fileName = txt;
                        string[] arrStr = txt.Split(',');

                        msg.Body.SetString("FILENAME", txt);
                        msg.Body.SetString("STATUS", msgg);
                        //msg.Body.SetString("MESSAGE", msgg);
                        producer.Send(msg);
                        Console.WriteLine("Message Send for " + txt);
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }




        public void postMessageforIssueCAMS(string txt, string msgg)
        {
            //Create the Connection factory
            try
            {
                string url = ConfigurationManager.AppSettings["queueURL"].ToString();
                string queueName = ConfigurationManager.AppSettings["issueCAMS"].ToString();

                IConnectionFactory factory = new ConnectionFactory(url);

                using (IConnection connection = factory.CreateConnection())
                using (ISession session = connection.CreateSession())
                {
                    IDestination destination = SessionUtil.GetDestination(session, queueName);
                    Console.WriteLine("Using destination: " + destination);

                    // Create a producer                
                    using (IMessageProducer producer = session.CreateProducer(destination))
                    {
                        // Start the connection so that messages will be processed.
                        connection.Start();
                        IMapMessage msg = producer.CreateMapMessage();
                        string fileName = txt;
                        string[] arrStr = txt.Split(',');

                        msg.Body.SetString("FILENAME", txt);
                        msg.Body.SetString("STATUS", msgg);
                        //msg.Body.SetString("MESSAGE", msgg);
                        producer.Send(msg);
                        Console.WriteLine("Message Send for " + txt);
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }

        public void postCAMSMessage(string folderCode, string msgg)
        {
            //Create the Connection factory
            try
            {
                string url = ConfigurationManager.AppSettings["queueURL"].ToString();
                string queueName = ConfigurationManager.AppSettings["queueNameCAMS"].ToString();

                IConnectionFactory factory = new ConnectionFactory(url);

                using (IConnection connection = factory.CreateConnection())
                using (ISession session = connection.CreateSession())
                {
                    IDestination destination = SessionUtil.GetDestination(session, queueName);
                    Console.WriteLine("Using destination: " + destination);

                    // Create a producer                
                    using (IMessageProducer producer = session.CreateProducer(destination))
                    {
                        // Start the connection so that messages will be processed.
                        connection.Start();
                        IMapMessage msg = producer.CreateMapMessage();
                        string fileName = folderCode;
                        msg.Body.SetString("FILENAME", folderCode);
                        msg.Body.SetString("STATUS", "supply_cams_bundle");
                        msg.Body.SetString("STAGE", "FV");
                        Console.WriteLine("System is in sleep mode for one minute after sending CAMS request to Dataset.");
                        System.Threading.Thread.Sleep(2 * 60 * 1000);
                        producer.Send(msg);                        
                        Console.WriteLine("System awake from sleep mode after two minute sending CAMS request to Dataset.");
                        Console.WriteLine("Message send for " + folderCode + " on Dataset queue");
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }

        public void postAMPDFMessage(string folderCode, string msgg)
        {
            //Create the Connection factory
            try
            {
                string url = ConfigurationManager.AppSettings["queueURL"].ToString();
                string queueName = ConfigurationManager.AppSettings["queueAMName"].ToString();

                IConnectionFactory factory = new ConnectionFactory(url);

                using (IConnection connection = factory.CreateConnection())
                using (ISession session = connection.CreateSession())
                {
                    IDestination destination = SessionUtil.GetDestination(session, queueName);
                    Console.WriteLine("Using destination: " + destination);
                    // Create a producer                
                    using (IMessageProducer producer = session.CreateProducer(destination))
                    {
                        // Start the connection so that messages will be processed.
                        connection.Start();
                        IMapMessage msg = producer.CreateMapMessage();
                        string fileName = folderCode;
                        msg.Body.SetString("FILENAME", "AM"+folderCode);
                        producer.Send(msg);
                    }
                }
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }
    }
}