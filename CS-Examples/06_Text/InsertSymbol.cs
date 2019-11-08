using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertSymbol
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Add a section.
            Section section = document.AddSection();

            //Add a paragraph.
            Paragraph paragraph = section.AddParagraph();

            //Use unicode characters to create symbol Ä.
            TextRange tr = paragraph.AppendText('\u00c4'.ToString());

            //Set the color of symbol Ä.
            tr.CharacterFormat.TextColor = Color.Red;

            //Add symbol Ë.
            paragraph.AppendText('\u00cb'.ToString());

            String result = "Result-InsertSymbol.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
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
