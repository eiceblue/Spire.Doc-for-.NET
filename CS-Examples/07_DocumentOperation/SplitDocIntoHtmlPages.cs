using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
namespace SplitDocIntoHtmlPages
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String input = @"..\..\..\..\..\..\..\Data\SplitDocIntoHtmlPages.doc";
            string outDir = Path.Combine("output");
            Directory.CreateDirectory(outDir);

            //Split a document into multiple html pages.
            SplitDocIntoMultipleHtml(input, outDir);
       }
        private void SplitDocIntoMultipleHtml(String input, string outDirectory)
        {
            Document document = new Document();
            document.LoadFromFile(input);

            Document subDoc = null;
            bool first = true;
            int index = 0;
            foreach (Section sec in document.Sections)
            {
                foreach (DocumentObject element in sec.Body.ChildObjects)
                {
                    if (IsInNextDocument(element))
                    {
                        if (!first)
                        {
                            //Embed css tyle and image data into html page
                            subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal;
                            subDoc.HtmlExportOptions.ImageEmbedded = true;
                            //Save to html file
                            subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index++)),FileFormat.Html);
                            subDoc = null;
                        }
                        first = false;
                    }
                    if (subDoc == null)
                    {
                        subDoc = new Document();
                        subDoc.AddSection();
                    }
                    subDoc.Sections[0].Body.ChildObjects.Add(element.Clone());
                }
            }
            if (subDoc != null)
            {
                //Embed css tyle and image data into html page
                subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal;
                subDoc.HtmlExportOptions.ImageEmbedded = true;
                //Save to html file
                subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index++)), FileFormat.Html);
            }
        }
        private bool IsInNextDocument(DocumentObject element)
        {
            if (element is Paragraph)
            {
                Paragraph p = element as Paragraph;
                if (p.StyleName == "Heading1")
                {
                    return true;
                }
            }
            return false;
        }
    }
}
