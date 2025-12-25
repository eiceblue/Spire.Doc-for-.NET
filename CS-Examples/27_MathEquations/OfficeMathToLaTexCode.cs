using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.OMath;
using System.IO;

namespace OfficeMathToLaTexCode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Load the existing Word document containing Office Math objects
            Document document = new Document(@"..\..\..\..\..\..\Data\OfficeMath.docx");

            // Create a StringBuilder to accumulate the LaTeX code strings
            StringBuilder stringBuilder = new StringBuilder();

            // Iterate through all sections in the document
            foreach (Section section in document.Sections)
            {
                // Iterate through all paragraphs within the current section's body
                foreach (Paragraph par in section.Body.Paragraphs)
                {
                    // Iterate through all child objects within the current paragraph
                    foreach (DocumentObject obj in par.ChildObjects)
                    {
                        // Attempt to cast the current object to an OfficeMath object
                        OfficeMath officeMath = obj as OfficeMath;

                        // If the cast fails (obj is not OfficeMath), skip to the next object
                        if (officeMath == null) continue;

                        // Convert the OfficeMath object to its LaTeX representation
                        string LaTexCode = officeMath.ToLaTexMathCode();

                        // Append the LaTeX code to the StringBuilder, followed by a new line
                        stringBuilder.AppendLine(LaTexCode);
                    }
                }
            }

            // Define the name of the output text file
            String outputFile = "OfficeMathToLaTexCode.txt";

            // Write the accumulated LaTeX code string to the output file
            File.WriteAllText(outputFile, stringBuilder.ToString());

            // Dispose of the Document object to release resources
            document.Dispose();

            WordDocViewer(outputFile);
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
