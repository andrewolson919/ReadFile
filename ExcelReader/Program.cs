using System;
using System.Threading;
using System.Text;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            DisplayEncodedText(new byte[] { 0x59, 0x6F, 0x75, 0x20, 0x6C, 0x69, 0x6B, 0x65, 0x20, 0x6D, 0x65, 0x6E });
            Thread.Sleep(10000);
        }

        static void DisplayEncodedText(byte[] encodedText)
        {
            Console.WriteLine(new UTF8Encoding().GetString(encodedText));
        }
    }
}
