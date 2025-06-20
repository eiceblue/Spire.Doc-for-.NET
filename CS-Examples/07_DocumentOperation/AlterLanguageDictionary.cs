using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AlterLanguageDictionary
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class.
			Document document = new Document();

			// Add a section to the document.
			Section sec = document.AddSection();

			// Add a paragraph to the section.
			Paragraph para = sec.AddParagraph();

			// Append text "corrige seg¨²n diccionario en ingl¨¦s" to the paragraph.
			TextRange txtRange = para.AppendText("corrige seg¨²n diccionario en ingl¨¦s");

			// Set the LocaleIdASCII property of the CharacterFormat for the text range to 10250.
			txtRange.CharacterFormat.LocaleIdASCII = 10250;

			// Specify the output file name.
			string result = "Result-AlterLanguageDictionary.docx";

			// Save the document to a file with the specified output file name and format (Docx2013).
			document.SaveToFile(result, FileFormat.Docx2013);

			// Clean up resources used by the document.
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
