using System;
using System.IO.Ports;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace SupplyTester_Interface
{
    class IIP_TESTER
    {
        //public const bool DEBUG = true;        // режим debug

        static SerialPort port = new SerialPort();
        static int portNumber = 0;
        static int row = 2;
        static int startID = 1;
        static int column = 2;
        static string path = Directory.GetCurrentDirectory() + @"\karta.xlsx";

        // Создаём экземпляр нашего приложения
        static Excel.Application excelApp = new Excel.Application();
        // Создаём экземпляр рабочий книги Excel
        static Excel.Workbook workBook;
        // Создаём экземпляр листа Excel
        static Excel.Worksheet workSheet;

        public static void Connect()
        {
            port.DataReceived += new SerialDataReceivedEventHandler(DataParsing);   // добавляем обработчик по приходу данных
            string[] availablePorts = SerialPort.GetPortNames();                    // сохраняем список доступных портов
            

            if (availablePorts.Count() == 0)                                        // проверка количества доступных портов
            {
                Console.WriteLine("Devices not detected...\nResearch? Y/N\n");
                string s = Console.ReadLine();
                if (s == "Y" || s == "y") availablePorts = SerialPort.GetPortNames();
                else Console.Clear();
            }
            else Console.WriteLine("Change device COM port number: \n");

            for (short i = 0; i < availablePorts.Count(); i++)                      // выводим список портов
            {
                Console.Write((i + 1) + " - ");
                Console.WriteLine(availablePorts[i]);
            }

            try
            {
                portNumber = int.Parse(Console.ReadLine());                         // вибираем порт
            }
            catch (Exception e)
            {
                //Console.WriteLine(e.ToString());              // вывод ошибки
                Console.WriteLine("\nChange device COM port number: \n");
                portNumber = int.Parse(Console.ReadLine());
                Console.Clear();
            }

            if (!port.IsOpen)                                   // настраиваем и открываем выбранный com порт
            {
                port.BaudRate = 9600;
                port.DtrEnable = true;
                port.RtsEnable = true;

                port.PortName = availablePorts[portNumber - 1].ToString();
                port.Open();
            }
            if (port.IsOpen)                                    // сообщаем статус подключения
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

        public static void OpenExcel()
        {
            workBook = excelApp.Workbooks.Open(path);
            workSheet = (Excel.Worksheet)workBook.Worksheets.Add();

            workSheet.Cells[1, 1] = "ID";
            workSheet.Cells[2, 1] = "1";

            excelApp.Visible = true;
            excelApp.UserControl = true;
        }

        /*public static int ChangeInstuction()
        {
            Console.WriteLine("Available instructions: ");
            Console.Write
            (
                "0 - Test C/V protection\n" + 
                "1 - Test voltage\n"    +
                "2 - Test current\n"    +
                "3 - Read voltage\n"    +
                "4 - Read current\n"
            );

            int command = int.Parse(Console.ReadLine());
            switch (command)
            {
                    case 0:
                        if (port.IsOpen)
                    {
                        Console.Clear();
                        port.Write("testall");
                    }
                    break;
                    case 1:
                        if (port.IsOpen)
                    {
                        Console.WriteLine();
                        Console.Write("Tested voltage: ");
                        port.Write("voltage");
                    }
                    break;
                    case 2:
                        if (port.IsOpen)
                    {
                        Console.WriteLine();
                        Console.Write("Tested current: ");
                        port.Write("current");
                    }
                    break;
                    case 3:
                        if (port.IsOpen)
                    {
                        Console.WriteLine();
                        Console.Write("Output voltage: ");
                        port.Write("readv");
                    }
                    break;
                    case 4:
                        if (port.IsOpen)
                    {
                        Console.WriteLine();
                        Console.Write("Output current: ");
                        port.Write("readc");
                    }
                    break;
            }

            return 0;
        }*/

        public static void DataParsing(object sender, SerialDataReceivedEventArgs e)
        {
            workSheet.Cells[row, column] = port.ReadLine().ToString();
            column++;
            if(column == 14)
            {
                row++;
                column = 2;
                workSheet.Cells[row, 1] = ++startID;
            }
        }
    }
}

