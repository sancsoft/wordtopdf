using System;
using System.Linq;
using WordToPDF.Library;

namespace WordToPDF.ConsoleApp
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Count() < 1)
            {
                Console.WriteLine("USAGE: WordToPDF <source> [destination]");
                return -1;
            }
            try
            {
                string inputFile = (args.Count() > 0) ? args[0] : "c:/users/mterry/git/wordtopdf/src/hello.docx";
                string outputFile = (args.Count() > 1) ? args[1] : null;
                ConvertService convertService = new ConvertService();
                convertService.Initialize();
                convertService.Convert(inputFile, ref outputFile);
                return 0;
            }
            catch(Exception e)
            {
                Console.WriteLine($"Failed to generate PDF - {e.Message}");
                return -1;
            }
        }
    }
}
