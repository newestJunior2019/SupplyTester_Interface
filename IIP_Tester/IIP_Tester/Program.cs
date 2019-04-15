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
            Console.Clear();

            IIP_TESTER.Connect();

            Console.ReadLine();
        }

    }
}

