using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CopyParagraph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object named document1
            Document document1 = new Document();

            // Load an existing document from the specified file path
            document1.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_5.docx");

            // Create a new Document object named document2
            Document document2 = new Document();

            // Get the first section of document1
            Section s = document1.Sections[0];

            // Get the first paragraph of section s
            Paragraph p1 = s.Paragraphs[0];

            // Get the second paragraph of section s
            Paragraph p2 = s.Paragraphs[1];

            // Add a new section to document2
            Section s2 = document2.AddSection();

            // Clone and add the cloned paragraph (NewPara1) from document1 to s2
            Paragraph NewPara1 = (Paragraph)p1.Clone();
            s2.Paragraphs.Add(NewPara1);

            // Clone and add the cloned paragraph (NewPara2) from document1 to s2
            Paragraph NewPara2 = (Paragraph)p2.Clone();
            s2.Paragraphs.Add(NewPara2);

            // Create a PictureWatermark object
            PictureWatermark WM = new PictureWatermark();

            // Set the Picture property of WM to an image from the specified file path
            WM.Picture = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.jpg");

            // Set the Watermark property of document2 to WM
            document2.Watermark = WM;

            // Specify the filename for the resulting document
            String result = "Result-CopyWordParagraph.docx";

            // Save document2 to the specified file in the Docx2013 format
            document2.SaveToFile(result, FileFormat.Docx2013);

            // Dispose of the resources used by document1 and document2
            document1.Dispose();
            document2.Dispose();

            //Launch the MS Word file.
            WordDocViewer(result);
        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
