using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
namespace CreateWordNote
{
    class Program
    {
        static void Main(string[] args)
        {
            Word.Application WordApp = new Word.Application();
            Word.Document WordDoc = WordApp.Documents.Add(Visible: true);
            WordDoc.Range(0, 0).Text = @"=Rand(1,2)";
            WordDoc.SaveAs2(FileName: @"C:\R\qu.docx");
            WordDoc.Close();
            WordApp.Quit();
            WordApp = null;
            WordDoc = null;

        }
    }
}
