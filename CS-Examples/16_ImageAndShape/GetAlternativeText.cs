using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;
using System.Text;
using Spire.Doc.Fields;

namespace GetAlternativeText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
			Document document = new Document();

			//Load the document from disk
			document.LoadFromFile(@"..\..\..\..\..\..\Data\ShapeWithAlternativeText.docx");

			//Create string builder
			StringBuilder builder = new StringBuilder();

			//Loop through shapes and get the AlternativeText
			foreach (Section section in document.Sections)
			{
				//Loop through the paragraphs in the section
				foreach (Paragraph para in section.Paragraphs)
				{
					//Loop through the child objects in the paragraph
					foreach (DocumentObject obj in para.ChildObjects)
					{
						//If the shape is a shape object
						if (obj is ShapeObject)
						{
							string text = (obj as ShapeObject).AlternativeText;
							//Append the alternative text in builder
							builder.AppendLine(text);
						}
					}
				}
			}

			//Write the content in txt file
			string result = "GetAlternativeText_result.txt";
			File.WriteAllText(result, builder.ToString());

			// Dispose the document
			document.Dispose();

            //Launch the file
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
