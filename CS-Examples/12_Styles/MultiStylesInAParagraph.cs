using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace MultiStylesInAParagraph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document
            Document doc = new Document();
            
            //Add a section
            Section section = doc.AddSection();

            //Add a paragraph
            Paragraph para = section.AddParagraph();

            //Add a text range 1 and set its style
            TextRange range = para.AppendText("Spire.Doc for .NET ");
            range.CharacterFormat.FontName = "Calibri";
            range.CharacterFormat.FontSize = 16f;
            range.CharacterFormat.TextColor = Color.Blue;
            range.CharacterFormat.Bold = true;
            range.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            //Add a text range 2 and set its style
            range = para.AppendText("is a professional Word .NET library");
            range.CharacterFormat.FontName = "Calibri";
            range.CharacterFormat.FontSize = 15f;

            //Save the Word document
            string output="MultiStylesInAParagraph_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

            //Launch the file
            FileViewer(output);
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
