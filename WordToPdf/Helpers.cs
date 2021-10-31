using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdf
{
    public static class Helpers
    {

        public static void WordToPdf(string FolderName, string InputDir, string ExportDir)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;

            DirectoryInfo dirInfo = new DirectoryInfo(FolderName);
            FileInfo[] wordFiles = dirInfo.GetFiles("*" + InputDir);

            word.Visible = false;
            word.ScreenUpdating = false;

            foreach (FileInfo wordFile in wordFiles)
            {
                Object filename = (Object)wordFile.FullName;

                Document doc = word.Documents.Open(ref filename);
                doc.Activate();

                object outputFileName = wordFile.FullName.Replace(InputDir, ExportDir);
                object fileFormat = WdSaveFormat.wdFormatPDF;

                doc.SaveAs(ref outputFileName,
                    ref fileFormat);

                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;
            }
           ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
        }
    }
}
