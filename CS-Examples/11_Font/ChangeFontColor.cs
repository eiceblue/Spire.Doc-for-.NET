using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ChangeFontColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the document
            string input = @"..\..\..\..\..\..\Data\Sample.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first section and first paragraph
            Section section = doc.Sections[0];
            Paragraph p1 = section.Paragraphs[0];

            //Iterate through the childObjects of the paragraph 1 
            foreach (DocumentObject childObj in p1.ChildObjects)
            {
                if (childObj is TextRange)
                {
                    //Change text color
                    TextRange tr = childObj as TextRange;
                    tr.CharacterFormat.TextColor = Color.RosyBrown;
                }
            }

            //Get the second paragraph
            Paragraph p2 = section.Paragraphs[1];

            //Iterate through the childObjects of the paragraph 2
            foreach (DocumentObject childObj in p2.ChildObjects)
            {
                if (childObj is TextRange)
                {
                    //Change text color
                    TextRange tr = childObj as TextRange;
                    tr.CharacterFormat.TextColor = Color.DarkGreen;
                }
            }

            //Save and launch document
            string output = "ChangeFontColor.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
