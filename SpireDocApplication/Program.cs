using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpireDocApplication
{
    class Program
    {
        static string path = @"D:\INGENIERIA\PROYECTOS\Info-Arch\Audi Report\Documents\";
        static string fileName = "DocTemplate.docx";

        static void Main(string[] args)
        {
            #region Console Title
            Console.WriteLine("\t\t\t\t\t =================================");
            Console.WriteLine("\t\t\t\t\t|------ Application Spire.Doc ----|");
            Console.WriteLine("\t\t\t\t\t =================================");
            #endregion

            Document document = new Document();


            Console.ReadKey();
        }
    }
}
