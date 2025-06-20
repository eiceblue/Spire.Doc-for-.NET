using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.OMath;

namespace OfficeMathToOfficeMathMLCode
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
            Document doc = new Document();
            // Load a Word document from a specific file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\ToOfficeMathMLCode.docx");
            // Create a StringBuilder to store the MathML code
            StringBuilder stringBuilder = new StringBuilder();
            // Iterate through sections in the document
            foreach (Section section in doc.Sections)
            {
                // Iterate through paragraphs in each section
                foreach (Paragraph par in section.Body.Paragraphs)
                {
                    // Iterate through child objects in each paragraph
                    foreach (DocumentObject obj in par.ChildObjects)
                    {
                        // Check if the object is an OfficeMath equation
                        OfficeMath omath = obj as OfficeMath;
                        if (omath == null) continue;
                        // Convert OfficeMath equation to MathML code
                        string mathml = omath.ToOfficeMathMLCode();
                        // Append MathML code to the StringBuilder
                        stringBuilder.Append(mathml);
                        stringBuilder.Append("\r\n");
                    }
                }
            }
            // Write the MathML code to a text file
            File.WriteAllText("OfficeMathToOfficeMathMLCode.txt", stringBuilder.ToString());     

            WordDocViewer("OfficeMathToOfficeMathMLCode.txt");
        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch(Exception e) {
                Debug.Write(e.StackTrace);
            }
        }

    }
}
