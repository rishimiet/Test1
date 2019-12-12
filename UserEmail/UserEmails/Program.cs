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

namespace UserEmails
{
    class Program
    {
        
        static void Main(string[] args)
        {
            GetUserEmail euMail = new GetUserEmail();
            euMail.ReceiveMails();
        }
        
    }
}
