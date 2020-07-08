
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Spire.Doc;
using Spire.Doc.Fields;
using System.Drawing;
using Spire.Doc.Documents;
using System.IO;

namespace pdfsToWordIndex
{
    class Program
    {


        static void Main(string[] args)
        {
            fromPdfDirectoryToWordFileWithHeadings();
        }

        public static void fromPdfDirectoryToWordFileWithHeadings()
        {
            DirectoryInfo directory = new DirectoryInfo(@"C:\Users\Lenovo\Desktop\transcripts");
            FileInfo[] files = directory.GetFiles();
            List<FileInfo> orderedList = files.OrderBy(x => x.Name).ToList();
            foreach (FileInfo file in orderedList)
            {
                writeFromPDFToWORD(file.FullName, file.Name);
            }
        }
        public static void writeFromPDFToWORD(string PDFpath, string name)
        {
            Document document = new Document();
            document.LoadFromFile(@"C:\Users\Lenovo\Desktop\index.docx", FileFormat.Docx); //no spaces in url
            Spire.Doc.Documents.Paragraph paraInserted = document.Sections[0].AddParagraph();

            TextRange textRange0 = paraInserted.AppendText("\n" + name + "\n");
            textRange0.CharacterFormat.TextColor = Color.Blue;
            textRange0.CharacterFormat.FontSize = 20;
            textRange0.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            TextRange textRange1 = paraInserted.AppendText(ExtractHeadlinesFromPdf(PDFpath));
            document.Sections[0].Paragraphs.Insert(0, document.Sections[0].Paragraphs[document.Sections[0].Paragraphs.Count - 1]);
            document.SaveToFile(@"C:\Users\Lenovo\Desktop\index.docx", FileFormat.Docx);
        }

        public static string ExtractHeadlinesFromPdf(string path)
        {
            PdfReader reader = new PdfReader(path);
            int intPageNum = reader.NumberOfPages;
            List<string> lines;
            string page;
            string specificLine = "";
            StringBuilder sb = new StringBuilder();

            for (int i = 1; i < intPageNum; i++)
            {
                page = PdfTextExtractor.GetTextFromPage(reader, i);

                lines = page.Split('\n').OfType<string>().ToList();

                foreach (string line in lines)
                {
                    if (line.Contains("Segment"))
                    {
                        specificLine = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(line));
                        sb.AppendLine(specificLine + "\n\r");
                    }
                }
            }
            return sb.ToString();
        }

    }

}
