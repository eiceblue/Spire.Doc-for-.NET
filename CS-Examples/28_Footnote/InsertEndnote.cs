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
            //Create a document and load file
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\InsertEndnote.doc");
            Section s = doc.Sections[0];
            Paragraph p = s.Paragraphs[1];

            //add endnote
            Footnote endnote = p.AppendFootnote(FootnoteType.Endnote);

            //append text
            TextRange text = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia");

            //set text format
            text.CharacterFormat.FontName = "Impact";
            text.CharacterFormat.FontSize = 14;
            text.CharacterFormat.TextColor = Color.DarkOrange;

            //Set marker format of endnote
            endnote.MarkerCharacterFormat.FontName = "Calibri";
            endnote.MarkerCharacterFormat.FontSize = 25;
            endnote.MarkerCharacterFormat.TextColor = Color.DarkBlue;

            //Save the document
            doc.SaveToFile("InsertEndnote.docx", FileFormat.Docx);

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
