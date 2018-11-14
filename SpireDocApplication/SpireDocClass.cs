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
    public class SpireDocClass
    {
        public Document document { get; set; }
        public string filePath { get; set; }
        public string fileName { get; set; }

        public SpireDocClass(string FilePath, string FileName)
        {
            filePath = FilePath;
            fileName = FileName;
            document = new Document(filePath + fileName);
        }

        public void ReplaceWithHTML(string html, string wordToReplace)
        {
            List<DocumentObject> replacement = new List<DocumentObject>();

            Section tempSection = document.AddSection();
            Paragraph paragraph = tempSection.AddParagraph();
            paragraph.AppendHTML(html);

            foreach (var obj in tempSection.Body.ChildObjects)
            {
                DocumentObject docObject = obj as DocumentObject;
                replacement.Add(docObject);
            }

            TextSelection[] selections = document.FindAllString(wordToReplace, false, true);
            List<TextRangeLocation> locations = new List<TextRangeLocation>();
            foreach (TextSelection selection in selections)
            {
                locations.Add(new TextRangeLocation(selection.GetAsOneRange()));
            }
            locations.Sort();
            foreach (TextRangeLocation location in locations)
            {
                ReplaceObj(location, replacement);
            }
            //remove the temp section
            document.Sections.Remove(tempSection);
            // Save Document
            document.SaveToFile(filePath + "DocGenerated.docx", FileFormat.Docx);

            Console.WriteLine("SUCCESS FULL..!!! Document Created.");
        }


        private void ReplaceObj(TextRangeLocation location, List<DocumentObject> replacement)
        {
            //will be replaced
            TextRange textRange = location.Text;
            //textRange index
            int index = location.Index;
            //owener paragraph
            Paragraph paragraph = location.Owner;
            //owner text body
            Body sectionBody = paragraph.OwnerTextBody;
            //get the index of paragraph in section
            int paragraphIndex = sectionBody.ChildObjects.IndexOf(paragraph);

            int replacementIndex = -1;
            if (index == 0)
            {
                //remove
                paragraph.ChildObjects.RemoveAt(0);
                replacementIndex = sectionBody.ChildObjects.IndexOf(paragraph);
            }
            else if (index == paragraph.ChildObjects.Count - 1)
            {
                paragraph.ChildObjects.RemoveAt(index);
                replacementIndex = paragraphIndex + 1;
            }
            else
            {
                //split owner paragraph
                Paragraph paragraph1 = (Paragraph)paragraph.Clone();
                while (paragraph.ChildObjects.Count > index)
                {
                    paragraph.ChildObjects.RemoveAt(index);
                }
                int i = 0;
                int count = index + 1;
                while (i < count)
                {
                    paragraph1.ChildObjects.RemoveAt(0);
                    i += 1;
                }
                sectionBody.ChildObjects.Insert(paragraphIndex + 1, paragraph1);

                replacementIndex = paragraphIndex + 1;
            }
            //insert replacement
            for (int i = 0; i <= replacement.Count - 1; i++)
            {
                sectionBody.ChildObjects.Insert(replacementIndex + i, replacement[i].Clone());
            }
        }
        public class TextRangeLocation : IComparable<TextRangeLocation>
        {
            public TextRangeLocation(TextRange text)
            {
                this.Text = text;
            }
            public TextRange Text
            {
                get { return m_Text; }
                set { m_Text = value; }
            }
            private TextRange m_Text;
            public Paragraph Owner
            {
                get { return this.Text.OwnerParagraph; }
            }
            public int Index
            {
                get { return this.Owner.ChildObjects.IndexOf(this.Text); }
            }
            public int CompareTo(TextRangeLocation other)
            {
                return -(this.Index - other.Index);
            }
        }
        public void ConvertToHtmlFile()
        {
            document.SaveToFile(filePath + "DocHTML\\DocumentTables.html", FileFormat.Html);
            Console.WriteLine("Saved as HTML File..!!!");
        }
    }
}
