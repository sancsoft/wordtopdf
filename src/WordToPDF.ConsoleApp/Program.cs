using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordToPDF.Library;

namespace WordToPDF.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = (args.Count() > 0) ? args[0] : "c:/users/mterry/git/wordtopdf/src/hello.docx";
            string outputFile = (args.Count() > 1) ? args[1] : null;
            ConvertService convertService = new ConvertService();
            convertService.Initialize();
            convertService.Convert(inputFile,outputFile);
        }
    }
}
