using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Formatting;
using Spire.Doc.Fields;

namespace SetFont
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			string input = @"..\..\..\..\..\..\Data\Sample.docx";

			//Create a Word document.
			Document doc = new Document();

			//Load the file from disk.
			doc.LoadFromFile(input);

			//Get the first section 
			Section s = doc.Sections[0];

			//Get the second paragraph
			Paragraph p = s.Paragraphs[1];

			//Create a characterFormat object
			CharacterFormat format = new CharacterFormat(doc);
			//Set font
			format.Font = new Font("Arial", 16);

			//Loop through the childObjects of paragraph 
			foreach (DocumentObject childObj in p.ChildObjects)
			{
				if (childObj is TextRange)
				{
					//Apply character format
					TextRange tr = childObj as TextRange;
					tr.ApplyCharacterFormat(format);
				}
			}

			//Save and launch document
			string output = "SetFont.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			//Dispose the Document
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
