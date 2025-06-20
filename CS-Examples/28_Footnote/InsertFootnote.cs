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
			// Create a new instance of Document
			Document document = new Document();

			// Load the Word document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\FootnoteExample.docx");

			// Find the specified string in the document
			TextSelection selection = document.FindString("Spire.Doc", false, true);

			// Get the selected text as a single range
			TextRange textRange = selection.GetAsOneRange();

			// Get the paragraph that contains the selected text
			Paragraph paragraph = textRange.OwnerParagraph;

			// Get the index of the selected text within the paragraph's child objects
			int index = paragraph.ChildObjects.IndexOf(textRange);

			// Append a footnote to the paragraph
			Footnote footnote = paragraph.AppendFootnote(FootnoteType.Footnote);

			// Insert the footnote into the paragraph's child objects at the specified index
			paragraph.ChildObjects.Insert(index + 1, footnote);

			// Add a paragraph to the footnote's text body and append text to it
			textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc");

			// Set the font name, size, and color for the appended text
			textRange.CharacterFormat.FontName = "Arial Black";
			textRange.CharacterFormat.FontSize = 10;
			textRange.CharacterFormat.TextColor = Color.DarkGray;

			// Set the font name, size, style, and color for the footnote marker
			footnote.MarkerCharacterFormat.FontName = "Calibri";
			footnote.MarkerCharacterFormat.FontSize = 12;
			footnote.MarkerCharacterFormat.Bold = true;
			footnote.MarkerCharacterFormat.TextColor = Color.DarkGreen;

			// Save the modified document to a file
			document.SaveToFile("AddFootnote.docx", FileFormat.Docx2010);

			// Dispose of the document object when finished using it
			document.Dispose();

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
