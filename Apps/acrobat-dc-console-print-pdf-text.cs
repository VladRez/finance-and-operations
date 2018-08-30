using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Acrobat;
namespace InvoiceToXLConsole
{
    class Program
    {
        static void Main(string[] args)
        {

            // Create an Acrobat Application object
            Type AcrobatAppType;
            AcrobatAppType = Type.GetTypeFromProgID("AcroExch.App");
            Acrobat.CAcroApp oAdobeApp = (Acrobat.CAcroApp)Activator.CreateInstance(AcrobatAppType);
            // Create an Acrobat Document object;
            Type AcrobatPDDocType;
            AcrobatPDDocType = Type.GetTypeFromProgID("AcroExch.PDDoc");
            Acrobat.CAcroPDDoc oAdobePDDoc = (Acrobat.CAcroPDDoc)Activator.CreateInstance(AcrobatPDDocType);
            // Create an Acrobat AV Document object;
            Type AcrobatAvType;
            AcrobatAvType = Type.GetTypeFromProgID("AcroExch.AVDoc");
            Acrobat.CAcroAVDoc oAdobeAVDoc = (Acrobat.CAcroAVDoc)Activator.CreateInstance(AcrobatAvType);

            oAdobePDDoc.Open(@"C:\dev\PALL.pdf");

            // Create the JavaScript object
            Object jsObj = oAdobePDDoc.GetJSObject();
            
            Type T = jsObj.GetType();

            for (int i = 97; i < 200; i++) {


                object[] getNWordParam = { 1, i, true };
                String word = (String)T.InvokeMember("getPageNthWord",
                               System.Reflection.BindingFlags.InvokeMethod |
                               System.Reflection.BindingFlags.Public |
                               System.Reflection.BindingFlags.Instance,
                               null, jsObj, getNWordParam);

                Console.WriteLine(word);
            }

            
            //Console.WriteLine("Total # of pages {0}", oAdobePDDoc.GetNumPages());
            //Int32 WordLimit = 50;
            
               
            



            oAdobePDDoc.Close();

        }
    }
}
