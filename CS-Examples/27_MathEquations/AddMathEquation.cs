using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields.OMath;
using Spire.Doc.Documents;

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
            string[] latexMathCode = { 
             @"x^{2}+\\sqrt{x^{2}+1}=2",
             @"2\alpha - \sin y + x",
             @"1 \over 2 + x", 
             @"(1 + \vert x-[a-b] \vert)",
             @"\mbox{if $x=1$ or $x=2$}",
             @"\begin{cases} 1 & \mbox{if $x>0$,} \\ 2 & \mbox{otherwise.} \end{cases}"        
                                     };
       
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\AddMathEquation.docx");
            Section section = doc.Sections[0];

            Paragraph paragraph;
            OfficeMath officeMath;
            //Add LaTeX code
            Table table1 = section.Tables[0] as Table;
            List<OfficeMath> mathEquations = new List<OfficeMath>();
            for (int i = 1; i < 7; i++)
            {
                paragraph = table1.Rows[i].Cells[0].AddParagraph();
                paragraph.Text = latexMathCode[i - 1];
                paragraph = table1.Rows[i].Cells[1].AddParagraph();
                officeMath = new OfficeMath(doc);
                officeMath.FromLatexMathCode(latexMathCode[i - 1]);
                paragraph.Items.Add(officeMath);
                mathEquations.Add(officeMath);
            }

            //Add MathML code
            Table table2 = section.Tables[1] as Table;
            for (int i = 1; i < 7; i++)
            {
                paragraph = table2.Rows[i].Cells[0].AddParagraph();
                paragraph.Text = mathEquations[i-1].ToMathMLCode();
                paragraph = table2.Rows[i].Cells[1].AddParagraph();
                officeMath = new OfficeMath(doc);
                officeMath.FromMathMLCode(mathEquations[i - 1].ToMathMLCode());
                paragraph.Items.Add(officeMath);

            }
            string result = "AddMathEquation_result.docx";
            doc.SaveToFile(result, FileFormat.Docx);
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
