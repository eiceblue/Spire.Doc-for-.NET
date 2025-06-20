using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddTabStopsToParagraph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new document
            Document document = new Document();

            // Add a section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section
            Paragraph paragraph1 = section.AddParagraph();

            // Add a tab stop at position 28 to the paragraph
            Tab tab = paragraph1.Format.Tabs.AddTab(28);
            tab.Justification = TabJustification.Left;

            // Append the text "Washing Machine" with a tab character
            paragraph1.AppendText("\tWashing Machine");

            // Add another tab stop at position 280 to the paragraph
            tab = paragraph1.Format.Tabs.AddTab(280);
            tab.Justification = TabJustification.Left;
            tab.TabLeader = TabLeader.Dotted;

            // Append the text "$650" with a tab character and dotted leader
            paragraph1.AppendText("\t$650");

            // Add a new paragraph to the section
            Paragraph paragraph2 = section.AddParagraph();

            // Add a tab stop at position 28 to the second paragraph
            tab = paragraph2.Format.Tabs.AddTab(28);
            tab.Justification = TabJustification.Left;

            // Append the text "Refrigerator" with a tab character
            paragraph2.AppendText("\tRefrigerator");

            // Add another tab stop at position 280 to the second paragraph
            tab = paragraph2.Format.Tabs.AddTab(280);
            tab.Justification = TabJustification.Left;
            tab.TabLeader = TabLeader.NoLeader;

            // Append the text "$800" with a tab character and no leader
            paragraph2.AppendText("\t$800");

            // Specify the filename for the resulting document
            string result = "Result-AddTabStopsToParagraph.docx";

            // Save the document to the specified file in the Docx2013 format
            document.SaveToFile(result, FileFormat.Docx2013);

            // Dispose of the document resources
            document.Dispose();

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
