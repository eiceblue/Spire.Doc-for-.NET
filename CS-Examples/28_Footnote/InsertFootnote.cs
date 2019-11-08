using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc.Fields;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertFootnote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\FootnoteExample.docx");

            //finds the first matched string.
            TextSelection selection = document.FindString("Spire.Doc", false, true);

            TextRange textRange = selection.GetAsOneRange();
            Paragraph paragraph = textRange.OwnerParagraph;
            int index = paragraph.ChildObjects.IndexOf(textRange);
            Footnote footnote = paragraph.AppendFootnote(FootnoteType.Footnote);
            paragraph.ChildObjects.Insert(index + 1, footnote);

            textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc");
            textRange.CharacterFormat.FontName = "Arial Black";
            textRange.CharacterFormat.FontSize = 10;
            textRange.CharacterFormat.TextColor = Color.DarkGray;

            footnote.MarkerCharacterFormat.FontName = "Calibri";
            footnote.MarkerCharacterFormat.FontSize = 12;
            footnote.MarkerCharacterFormat.Bold = true;
            footnote.MarkerCharacterFormat.TextColor = Color.DarkGreen;

            document.SaveToFile("AddFootnote.docx", FileFormat.Docx2010);

            //view the Word file.
            WordDocViewer("AddFootnote.docx");
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
