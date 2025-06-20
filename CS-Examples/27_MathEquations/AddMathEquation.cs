using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.OMath;

namespace AddMathEquation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define an array of LaTeX math code strings
            string[] latexMathCode = {
                @"x^{2}+\sqrt{x^{2}+1}=2",
                @"2\alpha - \sin y + x",
                @"1 \over 2 + x",
                @"(1 + \vert x-[a-b] \vert)",
                @"\mbox{if $x=1$ or $x=2$}",
                @"\begin{cases} 1 & \mbox{if $x>0$,} \ 2 & \mbox{otherwise.} \end{cases}"
            };

            // Create a new document object
            Document doc = new Document();

            // Load a document from the specified file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\AddMathEquation.docx");

            // Get the first section in the document
            Section section = doc.Sections[0];

            // Declare variables for paragraph and OfficeMath objects
            Paragraph paragraph;
            OfficeMath officeMath;

            // Get the first table in the section
            Table table1 = section.Tables[0] as Table;

            // Create a list to store the OfficeMath objects representing the math equations
            List<OfficeMath> mathEquations = new List <OfficeMath> ();

            // Iterate through the rows of the first table (excluding the header row)
            for (int i = 1; i < 7; i++)
            {
                // Get the first cell in the current row and add the LaTeX math code as text
                paragraph = table1.Rows[i].Cells[0].AddParagraph();
                paragraph.Text = latexMathCode[i - 1];

                // Get the second cell in the current row and create an OfficeMath object from the LaTeX math code
                paragraph = table1.Rows[i].Cells[1].AddParagraph();
                officeMath = new OfficeMath(doc);
                officeMath.FromLatexMathCode(latexMathCode[i - 1]);
                paragraph.Items.Add(officeMath);

                // Add the OfficeMath object to the list
                mathEquations.Add(officeMath);
            }

            // Get the second table in the section
            Table table2 = section.Tables[1] as Table;

            // Iterate through the rows of the second table (excluding the header row)
            for (int i = 1; i < 7; i++)
            {
                // Get the first cell in the current row and add the MathML code of the corresponding OfficeMath object as text
                paragraph = table2.Rows[i].Cells[0].AddParagraph();
                paragraph.Text = mathEquations[i - 1].ToMathMLCode();

                // Get the second cell in the current row and create an OfficeMath object from the MathML code
                paragraph = table2.Rows[i].Cells[1].AddParagraph();
                officeMath = new OfficeMath(doc);
                officeMath.FromMathMLCode(mathEquations[i - 1].ToMathMLCode());
                paragraph.Items.Add(officeMath);
            }

            // Specify the output file path
            string result = "AddMathEquation_result.docx";

            // Save the modified document to the output file in DOCX format
            doc.SaveToFile(result, FileFormat.Docx);

            // Dispose the document object
            doc.Dispose();
			
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
