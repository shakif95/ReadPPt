using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadPPtFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            String path = AppDomain.CurrentDomain.BaseDirectory + "AWS_CloudFormation.pptx";

            Console.Write(Read.Start(path));
            Console.ReadKey();
        }
    }
}
