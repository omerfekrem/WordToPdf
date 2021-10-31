using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Helpers.WordToPdf("WordAndPdf", ".doc", ".pdf");
        }
    }
}
