using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetIndentByCharacter
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
   
            // Create a new Document object
            Document document = new Document();

            // Add a section to the document
            Section sec = document.AddSection();

            // Add a paragraph for the title
            Paragraph para = sec.AddParagraph();
            para.AppendText("Paragraph Formatting");
            para.ApplyStyle(BuiltinStyle.Title);

            // Add a paragraph with indent settings
            para = sec.AddParagraph();
            para.AppendText( "This paragraph is indent as follows: Indent 2 characters on the left and 5 characters on the right.");
            para.Format.LeftIndentChars= 2f;
            para.Format.RightIndentChars = 5f;

            // Specify the output file name for the modified document
            string output = "SetIndentByCharacter_output.docx";

           // Save the modified document to the specified file format
           document.SaveToFile(output, FileFormat.Docx);

           // Dispose the Document object to free resources
          document.Dispose();
			
            WordDocViewer(output);
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
