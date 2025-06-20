using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Documents;

namespace InsertEndnote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Create a new document object
            Document doc = new Document();

            // Load a document from the specified file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\InsertEndnote.doc");

            // Get the first section in the document
            Section s = doc.Sections[0];

            // Get the second paragraph in the section (index 1)
            Paragraph p = s.Paragraphs[1];

            // Append an endnote to the paragraph
            Footnote endnote = p.AppendFootnote(FootnoteType.Endnote);

            // Add a paragraph to the endnote's text body and append the reference text
            TextRange text = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia");

            // Set the font name, size, and text color of the reference text
            text.CharacterFormat.FontName = "Impact";
            text.CharacterFormat.FontSize = 14;
            text.CharacterFormat.TextColor = Color.DarkOrange;

            // Set the font name, size, and text color of the endnote marker
            endnote.MarkerCharacterFormat.FontName = "Calibri";
            endnote.MarkerCharacterFormat.FontSize = 25;
            endnote.MarkerCharacterFormat.TextColor = Color.DarkBlue;

            // Save the modified document to the output file in DOCX format
            doc.SaveToFile("InsertEndnote.docx", FileFormat.Docx);

            // Dispose the document object
            doc.Dispose();

            //Launch the Word file
            WordDocViewer("InsertEndnote.docx");

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
