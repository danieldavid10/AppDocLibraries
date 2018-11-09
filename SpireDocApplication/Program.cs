using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpireDocApplication
{
    class Program
    {
        static string fileName = "DocTemplate.docx";
        static string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Documents\");

        static void Main(string[] args)
        {
            #region Console Title
            Console.WriteLine("\t\t\t\t\t =================================");
            Console.WriteLine("\t\t\t\t\t|------ Application Spire.Doc ----|");
            Console.WriteLine("\t\t\t\t\t =================================");
            Console.WriteLine("Working with the Document...");
            #endregion
            SpireDocClass spireDoc = new SpireDocClass(filePath, fileName);

            spireDoc.ReplaceWithHTML("<p style='text-align: center; '><b>AUDI REPORT PROTOTIPE</b></p>", "{{ProjectName}}");

            SpireDocClass spireDoc2 = new SpireDocClass(filePath, "DocGenerated.docx");
            spireDoc2.ConvertToHtmlFile();

            Console.ReadKey();
        }
    }
}


