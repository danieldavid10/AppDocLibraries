using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlAgilyPackApplication
{
    class Program
    {
        static string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Documents\HTML\");
        static void Main(string[] args)
        {
            #region Console Title
            Console.WriteLine("\t\t\t\t\t =========================================");
            Console.WriteLine("\t\t\t\t\t|------ Application Html Agility Pack ----|");
            Console.WriteLine("\t\t\t\t\t =========================================");
            Console.WriteLine("Working with the Document...");
            #endregion

            var doc = new HtmlDocument();
            doc.Load(filePath + "GoogleDocument.html");

            var htmlBody = doc.DocumentNode.SelectSingleNode("//div");

            HtmlNodeCollection childNodes = htmlBody.ChildNodes;

            foreach (var node in childNodes)
            {
                if (node.InnerText == "Project")
                {
                    Console.WriteLine(node.OuterHtml + "\n");
                }
            }

            Console.ReadKey();
        }
    }
}
