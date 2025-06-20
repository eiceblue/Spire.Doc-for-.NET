using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateBarcode
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
			Document doc = new Document();

			// Add a new section to the document and get its first paragraph.
			Paragraph p = doc.AddSection().AddParagraph();

			// Append the text "H63TWX11072" to the paragraph and obtain the TextRange object.
			TextRange txtRang = p.AppendText("H63TWX11072");

			// Set the font name for the text range to "C39HrP60DlTt".
			txtRang.CharacterFormat.FontName = "C39HrP60DlTt";

			// Set the font size for the text range to 80.
			txtRang.CharacterFormat.FontSize = 80;

			// Set the text color for the text range to SeaGreen.
			txtRang.CharacterFormat.TextColor = Color.SeaGreen;

			// Specify the file name for the resulting document.
			string output = "CreateBarcode.docx";

			// Save the document to a file with the specified file name and format (Docx).
			doc.SaveToFile(output, FileFormat.Docx);

			// Clean up resources used by the document.
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
