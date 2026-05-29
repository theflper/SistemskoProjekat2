using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Threading.Tasks;
using ExcelDataReader;

namespace SistemskoProjekat1
{
    internal class Program
    {
        //promeni arhitekturu imam while petlji prihvatam zahteve
        static void Main(string[] args)
        {
            Server s = new Server();
            //da ne bi blokirali unos sa Strart jer on ima blokirjuce funkcije u sebi
            Thread t = new Thread(() => s.Start());
            t.IsBackground = true;
            t.Start();
            string str;
            while (true)
            {
                str = Console.ReadLine();
                if (str == "stop")
                {
                    s.Stop();
                    break;
                }
            }
        }
    }
}
