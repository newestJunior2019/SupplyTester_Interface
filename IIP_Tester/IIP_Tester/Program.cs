using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Net.Sockets;
using System.Management;
using Microsoft.Win32;

namespace SupplyTester_Interface
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;
            ConsoleCancelEventHandler Closed = new ConsoleCancelEventHandler(ClosingEvent);
            Console.Clear();

            IIP_TESTER.Connect();
            IIP_TESTER.ChangeInstuction();

            Console.ReadLine();
            
        }
        
        static void ClosingEvent(object sender, ConsoleCancelEventArgs e)
        {
            IIP_TESTER.Disconnect();
        }
    }
}

