using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;
using System.Text;

namespace GetDocumentProperties
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

			// Load a Word document from a specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Properties.docx");

			// Create a StringBuilder to store the content
			StringBuilder content = new StringBuilder();

			// Retrieve built-in document properties
			string title = document.BuiltinDocumentProperties.Title;
			string comments = document.BuiltinDocumentProperties.Comments;
			string author = document.BuiltinDocumentProperties.Author;
			string keywords = document.BuiltinDocumentProperties.Keywords;
			string company = document.BuiltinDocumentProperties.Company;

			// Format the built-in document properties into a string
			string result = string.Format("The Builtin document properties:\r\nTitle: " + title + ".\r\nComments: " + comments + ".\r\nAuthor: " + author + ".\r\nKeywords: " + keywords + ".\r\nCompany: " + company);

			// Append the result to the content StringBuilder
			content.AppendLine(result + "\r\nThe custom document properties:");

			// Iterate through each custom document property and append it to the content StringBuilder
			for (int i = 0; i < document.CustomDocumentProperties.Count; i++)
			{
				string customProperties = string.Format(document.CustomDocumentProperties[i].Name + ": " + document.CustomDocumentProperties[i].Value);
				content.AppendLine(customProperties);
			}

			// Write the content to a text file named "Output.txt"
			File.WriteAllText("Output.txt", content.ToString());

			// Dispose of the Document object to free up resources
			document.Dispose();

            //Launch the txt file.
            WordDocViewer("Output.txt");
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
