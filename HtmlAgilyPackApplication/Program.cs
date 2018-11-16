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

            string text = getHtmlContent(childNodes, "Objective:", "Background:");

            Console.ReadKey();
        }

        private static string getHtmlContent(HtmlNodeCollection childNodes, string iniText, string endText)
        {
            int ini = getIndexNode(childNodes, iniText) + 1;
            int end = getIndexNode(childNodes, endText);
            if (ini == 0 || end == 0)
            {
                throw new Exception("Range Index Error");
            }
            else
            {
                HtmlNode divNode = HtmlNode.CreateNode("<div id='ProjectObjetive '></div>");
                for (int i = ini; i < end; i++)
                {
                    divNode.AppendChild(childNodes[i]);
                }
                return divNode.OuterHtml;
            }
        }

        private static int getIndexNode(HtmlNodeCollection childNodes, string text)
        {
            for (int i = 0; i < childNodes.Count; i++)
            {
                if (childNodes[i].InnerText.Contains(text))
                {
                    return i;
                }
            }
            return 0;
        }

        private static void SaveHtml(HtmlDocument document)
        {
            FileStream sw = new FileStream("FileModificated.html", FileMode.Create);
            document.Save(sw);
            sw.Close();
        }
    }
}
