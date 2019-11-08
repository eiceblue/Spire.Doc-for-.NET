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
using System.IO;
namespace GetMathEquation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\GetMathEquation.docx");
            List<OfficeMath> mathEquations = new List<OfficeMath>();
            StringBuilder stringBuilder = new StringBuilder();
            foreach (Section section in doc.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    foreach (DocumentObject obj in paragraph.ChildObjects)
                    {
                        if (obj is OfficeMath)
                        {
                            stringBuilder.AppendLine((obj as OfficeMath).ToMathMLCode());
                            stringBuilder.AppendLine();
                            mathEquations.Add(obj as OfficeMath);
                        }
                    }

                }
            }
            string output ="MathMLCode.txt";
            File.WriteAllText(output, stringBuilder.ToString());
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
