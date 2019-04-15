using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;

namespace SupplyTester_Interface
{
    class IIP_TESTER
    {
        static SerialPort port = new SerialPort();
        static int portNumber = 0;

        public static void Connect()
        {
            port.DataReceived += new SerialDataReceivedEventHandler(DataParsing);   // добавляем обработчик по приходу данных
            string[] availablePorts = SerialPort.GetPortNames();                    // сохраняем список доступных портов
            if (availablePorts.Count() == 0)
            {
                Console.WriteLine("Devices not detected...\nResearch? Y/N\n");
                string s = Console.ReadLine();
                if (s == "Y" || s == "y") availablePorts = SerialPort.GetPortNames();
                else Console.Clear();
            }
            else Console.WriteLine("Change device COM port number: \n");

            for (short i = 0; i < availablePorts.Count(); i++)                       // выводим список портов
            {
                Console.Write((i + 1) + " - ");
                Console.WriteLine(availablePorts[i]);
            }

            try
            {
                portNumber = int.Parse(Console.ReadLine());         // вибираем порт
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.WriteLine("\nChange device COM port number: \n");
                portNumber = int.Parse(Console.ReadLine());
                Console.Clear();
            }

            if (!port.IsOpen)
            {
                port.BaudRate = 38400;
                port.DtrEnable = true;
                port.RtsEnable = true;

                port.PortName = availablePorts[portNumber - 1].ToString();
                port.Open();
            }
            if (port.IsOpen)
            {
                Console.Clear();
                Console.WriteLine("Connected to " + availablePorts[portNumber - 1].ToString());
            }
            else Console.WriteLine("Connection ERROR!");
        }

        public static void Disconnect()
        {
            if (port.IsOpen)
            {
                port.Close();
                if (!port.IsOpen) Console.WriteLine("Disconnected!");
                else Console.WriteLine("Disconnection Error!");
            }
        }

        public static int ChangeInstuction()
        {
            Console.WriteLine("Available instructions: ");
            Console.Write
                (
                    "1 - Test voltage\n" +
                    "2 - Test current\n" +
                    "3 - Read voltage\n" +
                    "4 - Read current\n"
                );

            int command = int.Parse(Console.ReadLine());
            switch (command)
            {
                case 1:
                    if (port.IsOpen)
                    {
                        port.Write("voltage");
                    }
                    break;
                case 2:
                    if (port.IsOpen)
                    {
                        port.Write("current");
                    }
                    break;
                case 3:
                    if (port.IsOpen)
                    {
                        port.Write("readv");
                    }
                    break;
                case 4:
                    if (port.IsOpen)
                    {
                        port.Write("readc");
                    }
                    break;
            }

            return 0;
        }

        public static void DataParsing(object sender, SerialDataReceivedEventArgs e)
        {
            Console.WriteLine(port.ReadLine().ToString());
        }
    }
}

