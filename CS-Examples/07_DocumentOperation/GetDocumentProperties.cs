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
            //Create a document
            Document document = new Document();

            //Load the document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Properties.docx");

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Get Builtin document properties
            string title = document.BuiltinDocumentProperties.Title;
            string comments = document.BuiltinDocumentProperties.Comments;
            string author = document.BuiltinDocumentProperties.Author;
            string keywords = document.BuiltinDocumentProperties.Keywords;
            string company = document.BuiltinDocumentProperties.Company;

            //Set string format for displaying
            string result = string.Format("The Builtin document properties:\r\nTitle: " + title + ".\r\nComments: " + comments + ".\r\nAuthor: " + author + ".\r\nKeywords: " + keywords + ".\r\nCompany: " + company);

            //Add result string to StringBuilder
            content.AppendLine(result + "\r\nThe custom document properties:");

            //Get custom document properties
            for (int i = 0; i < document.CustomDocumentProperties.Count; i++)
            {
                string customProperties = string.Format(document.CustomDocumentProperties[i].Name + ": " + document.CustomDocumentProperties[i].Value);
                content.AppendLine(customProperties);
            }

            //Save them to a txt file
            File.WriteAllText("Output.txt", content.ToString());

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
