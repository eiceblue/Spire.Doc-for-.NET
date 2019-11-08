using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddPageNumbersInSections
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_4.docx");

            //Repeat step2 and Step3 for the rest sections, so change the code with for loop.
            for (int i = 0; i < 3; i++)
            {
                HeaderFooter footer = document.Sections[i].HeadersFooters.Footer;
                Paragraph footerParagraph = footer.AddParagraph();
                footerParagraph.AppendField("page number", FieldType.FieldPage);
                footerParagraph.AppendText(" of ");
                footerParagraph.AppendField("number of pages", FieldType.FieldSectionPages);
                footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

                if (i == 2)
                    break;
                else
                {
                    document.Sections[i + 1].PageSetup.RestartPageNumbering = true;
                    document.Sections[i + 1].PageSetup.PageStartingNumber = 1;
                }
            }

            String result = "Result-AddPageNumbersInSections.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the Ms Word file.
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
